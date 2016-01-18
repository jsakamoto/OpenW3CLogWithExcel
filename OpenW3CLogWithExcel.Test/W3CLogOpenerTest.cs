using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Xunit;
using Moq;

namespace OpenW3CLogWithExcel.Test
{
    public class W3CLogOpenerTest
    {
        private string PathOf(string fileName) => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", fileName);

        [Fact]
        public void Open_EmptyFile_Test()
        {
            var emptyFilePath = PathOf("size-zero.log");
            var mockShell = new Mock<IShell>();
            mockShell.Setup(m => m.Open(emptyFilePath)).Returns(default(Process));

            var opener = new W3CLogOpener(mockShell.Object);
            opener.Open(emptyFilePath);

            mockShell.VerifyAll();
        }

        [Fact]
        public void Open_LoremTextFile_Test()
        {
            var loremFilePath = PathOf("lorem.log");
            var mockShell = new Mock<IShell>();
            mockShell.Setup(m => m.Open(loremFilePath)).Returns(default(Process));

            var opener = new W3CLogOpener(mockShell.Object);
            opener.Open(loremFilePath);

            mockShell.VerifyAll();
        }

        [Fact]
        public void Open_W3CLogTextFile_Test()
        {
            var xlsxPath = default(string);
            var mockShell = new Mock<IShell>();
            mockShell
                .Setup(m => m.Open(It.Is<string>(v => v.EndsWith(".xlsx"))))
                .Callback<string>(path =>
                {
                    xlsxPath = path;
                    // .xlsx file exactly exists.
                    File.Exists(xlsxPath).IsTrue();
                    // .xlsx file is read only.
                    File.GetAttributes(xlsxPath).HasFlag(FileAttributes.ReadOnly).IsTrue();
                })
                .Returns(() => Process.Start("rundll32.exe"));

            var opener = new W3CLogOpener(mockShell.Object);
            var w3cLogFilePath = PathOf("w3c.log");
            opener.Open(w3cLogFilePath);

            mockShell.VerifyAll();
            // Temporary .xlsx file was sweeped.
            File.Exists(xlsxPath).IsFalse();
        }
        [Fact]
        public void Open_W3CLogTextFile_ExcelProcess_is_Null_Test()
        {
            var xlsxPath = default(string);
            var mockShell = new Mock<IShell>();
            mockShell
                .Setup(m => m.Open(It.Is<string>(v => v.EndsWith(".xlsx"))))
                .Callback<string>(path =>
                {
                    xlsxPath = path;
                    // .xlsx file exactly exists.
                    File.Exists(xlsxPath).IsTrue();
                    // .xlsx file is read only.
                    File.GetAttributes(xlsxPath).HasFlag(FileAttributes.ReadOnly).IsTrue();
                })
                // ! Retun Null when already Excel instances are there.
                .Returns(default(Process));

            var opener = new W3CLogOpener(mockShell.Object);
            var w3cLogFilePath = PathOf("w3c.log");
            opener.Open(w3cLogFilePath);

            mockShell.VerifyAll();
            // Temporary .xlsx file was sweeped.
            File.Exists(xlsxPath).IsFalse();
        }
    }
}
