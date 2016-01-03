using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OpenW3CLogWithExcel
{
    internal static class Shell
    {
        public static Process Open(string path)
        {
            var pi = new ProcessStartInfo
            {
                FileName = path,
                Verb = "Open",
                UseShellExecute = true
            };
            return Process.Start(pi);
        }
    }
}
