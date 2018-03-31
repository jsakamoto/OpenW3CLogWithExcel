using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace OpenW3CLogWithExcel
{
    public class W3CLogOpener
    {
        public event EventHandler Exit;

        public event EventHandler Converted;

        public event EventHandler<ProgressEventArgs> Progress;

        private IShell Shell;

        public W3CLogOpener()
        {
            this.Shell = new Shell();
        }

        public W3CLogOpener(IShell shell)
        {
            this.Shell = shell;
        }

        public void Open(string path)
        {
            // Start reading file...
            var lines = ReadLines(path);

            // Find first "#Fields:..." row.
            var marker = lines
                .SkipWhile(line => line.Text.StartsWith("#") && !line.Text.StartsWith("#Fields:"))
                .FirstOrDefault();

            // if not found, it's not W3C log file, then open the path with "Open" verb and exit.
            if ((marker.Text?.StartsWith("#Fields:") ?? false) == false)
            {
                Shell.Open(path);
                Exit?.Invoke(this, EventArgs.Empty);
                return;
            }

            // Generate .xlsx path as temporary file.
            var tmpDirPath = EnumGuidStrings()
                .Select(guid => Path.Combine(Path.GetTempPath(), guid))
                .First(dirPath => !Directory.Exists(dirPath) && !File.Exists(dirPath));
            Directory.CreateDirectory(tmpDirPath);
            var xlsxPath = Path.Combine(tmpDirPath, $"{Path.GetFileNameWithoutExtension(path)}.xlsx");

            try
            {
                // Convert from W3C text log file to .xlsx file.
                ConvertW3CLogToXlsx(lines, marker, xlsxPath);
                Converted?.Invoke(this, EventArgs.Empty);

                // Open with "Open" verb for .xlsx file.
                Shell.Open(xlsxPath)?.WaitForExit();
            }
            finally
            {
                // Sweep .xlsx as temporary file.
                try
                {
                    if (File.Exists(xlsxPath))
                    {
                        File.SetAttributes(xlsxPath, FileAttributes.Normal);
                        File.Delete(xlsxPath);
                    }
                }
                catch { }
                try { if (Directory.Exists(tmpDirPath)) Directory.Delete(tmpDirPath); } catch { }
            }

            Exit?.Invoke(this, EventArgs.Empty);
        }

        private void ConvertW3CLogToXlsx(IEnumerable<Line> lines, Line marker, string xlsxPath)
        {
            using (var xlbook = new XLWorkbook())
            using (var xlsheet = xlbook.AddWorksheet("Sheet1"))
            {
                // Query columns header texts and adjust date & time column.
                var headers = marker.Text.Split(' ').Skip(1).ToList();
                var colIndexOfDate = headers.IndexOf("date");
                var colIndexOfTime = headers.IndexOf("time");
                if (colIndexOfDate != -1 && colIndexOfTime != -1)
                {
                    headers[colIndexOfDate] = "date-time";
                    headers.RemoveAt(colIndexOfTime);
                    xlsheet.Column(colIndexOfDate + 1).Style.NumberFormat.Format = "yyyy/mm/dd hh:mm:ss";
                }

                // Build column header on .xlsx
                using (var headerRow = xlsheet.Row(1))
                    headers.Each((header, index) =>
                    {
                        headerRow.Cell(index + 1).Value = header;
                    });
                xlsheet.SheetView.FreezeRows(1);
                xlsheet.RangeUsed().SetAutoFilter();

                // Enumerate all content lines, and write into .xlsx
                var prevProgress = 0;
                lines
                .Where(line => !line.Text.StartsWith("#"))
                .Each((line, lineIndex) =>
                {
                    var newProgress = (int)Math.Ceiling(line.Progress * 100);
                    if (prevProgress != newProgress)
                    {
                        prevProgress = newProgress;
                        Progress?.Invoke(this, new ProgressEventArgs(newProgress));
                    }

                    // Adjust date & time column.
                    var values = line.Text.Split(' ').ToList();
                    if (colIndexOfDate != -1 && colIndexOfTime != -1)
                    {
                        values[colIndexOfDate] += " " + values[colIndexOfTime];
                        values.RemoveAt(colIndexOfTime);
                    }

                    using (var valueRow = xlsheet.Row(lineIndex + 1 + 1))
                        values.Each((value, colIndex) =>
                        {
                            valueRow.Cell(colIndex + 1).Value = value;
                        });
                });

                xlsheet.Columns().AdjustToContents(minWidth: 0, maxWidth: 100);

                xlbook.SaveAs(xlsxPath);
            }

            File.SetAttributes(xlsxPath, FileAttributes.ReadOnly);
        }

        private struct Line
        {
            public string Text { get; set; }

            public double Progress { get; set; }
        }

        private static IEnumerable<Line> ReadLines(string path)
        {
            using (var s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1024))
            using (var r = new StreamReader(s, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024))
            {
                var fileSize = s.Length;
                while (!r.EndOfStream)
                {
                    var line = new Line
                    {
                        Text = r.ReadLine(),
                        Progress = (double)s.Position / fileSize
                    };
                    Debug.WriteLine($"size:{fileSize} / read:{s.Position} / progress:{Math.Ceiling(line.Progress * 100)}");
                    yield return line;
                }
            }
        }

        private static IEnumerable<string> EnumGuidStrings()
        {
            for (; ; ) yield return Guid.NewGuid().ToString("N");
        }
    }
}
