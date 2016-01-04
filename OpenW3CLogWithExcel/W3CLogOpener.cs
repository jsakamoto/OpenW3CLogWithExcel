using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace OpenW3CLogWithExcel
{
    public class W3CLogOpener
    {
        public void Open(string path)
        {
            // Start reading file...
            var lines = ReadLines(path);

            // Find first "#Fields:..." row.
            var marker = lines
                .SkipWhile(line => line.StartsWith("#") && !line.StartsWith("#Fields:"))
                .FirstOrDefault();

            // if not found, it's not W3C log file, then open the path with "Open" verb and exit.
            if (marker?.StartsWith("#Fields:") == false)
            {
                Shell.Open(path);
                return;
            }

            // Generate .xlsx path as temporary file.
            var xlsxPath = Path.Combine(Path.GetTempPath(), $"~{Guid.NewGuid():N}.xlsx");

            // Convert from W3C text log file to .xlsx file.
            using (var xlbook = new XLWorkbook())
            using (var xlsheet = xlbook.AddWorksheet("Sheet1"))
            {
                // Query columns header texts and adjust date & time column.
                var headers = marker.Split(' ').Skip(1).ToList();
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
                lines
                .Where(line => !line.StartsWith("#"))
                .Each((line, lineIndex) =>
                {
                    // Adjust date & time column.
                    var values = line.Split(' ').ToList();
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

            // Open with "Open" verb for .xlsx file.
            File.SetAttributes(xlsxPath, FileAttributes.ReadOnly);
            var xlapp = Shell.Open(xlsxPath);
            xlapp.WaitForExit();

            // Sweep .xlsx as temporary file.
            try
            {
                File.SetAttributes(xlsxPath, FileAttributes.Normal);
                File.Delete(xlsxPath);
            }
            catch (Exception) { }
        }

        private static IEnumerable<string> ReadLines(string path)
        {
            using (var r = new StreamReader(path, detectEncodingFromByteOrderMarks: true))
            {
                while (!r.EndOfStream)
                {
                    yield return r.ReadLine();
                }
            }
        }
    }
}
