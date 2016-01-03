using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OpenW3CLogWithExcel.Installer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            const string verb = "openw3clogwithexcel";
            var verbPath = $@"txtfile\shell\{verb}";
            var HKCR = Registry.ClassesRoot;
            var shellKey = HKCR.OpenSubKey(verbPath);
            if (shellKey == null)
            {
                var commandKey = HKCR.OpenSubKey(@"0ea92b0253b34c489be1c97\shell\open\command");
                var commandValue = commandKey?.GetValue("")?.ToString();
                if (commandValue == null) return;

                shellKey = HKCR.CreateSubKey(verbPath);
                shellKey.SetValue("", "Open W3C log with Excel");

                var openLogCommanKey = shellKey.CreateSubKey("command");
                openLogCommanKey.SetValue("", commandValue);

                HKCR.OpenSubKey(@"txtfile\shell", writable: true).SetValue("", verb);
            }
        }
    }
}
