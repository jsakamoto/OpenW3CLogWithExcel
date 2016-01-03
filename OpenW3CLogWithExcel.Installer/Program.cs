using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.Reflection;

namespace OpenW3CLogWithExcel.Installer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            switch (args.FirstOrDefault())
            {
                case "install":
                    Install();
                    break;
                case "uninstall":
                    Uninstall();
                    break;
                default:
                    break;
            }
        }

        private const string verb = "openw3clogwithexcel";

        private static string verbPath => $@"txtfile\shell\{verb}";

        /// <summary>
        /// Execute install prcoess.
        /// </summary>
        private static void Install()
        {
            // Check installed or not: if exists "HKCR\txtfile\shell\openw3clogwithexcel", nothing to do.
            var HKCR = Registry.ClassesRoot;
            var shellKey = HKCR.OpenSubKey(verbPath);
            if (shellKey != null) return;

            // Add "Open W3C log with Excel" verb by cloning command which ClickOnce created.
            var commandKey = HKCR.OpenSubKey(@"0ea92b0253b34c489be1c97\shell\open\command");
            var commandValue = commandKey?.GetValue("")?.ToString();
            if (commandValue != null)
            {
                shellKey = HKCR.CreateSubKey(verbPath);
                shellKey.SetValue("", "Open W3C log with Excel");

                var openLogCommanKey = shellKey.CreateSubKey("command");
                openLogCommanKey.SetValue("", commandValue);

                // Set "Open W3C log with Excel" to default verb.
                HKCR.OpenSubKey(@"txtfile\shell", writable: true).SetValue("", verb);
            }

            // Override uninstall command, from ClickOnce original uninstaller to this custom uninstall process.
            var uninstallTheAppKey = FindUninstallTheAppRegKey(writable: true);
            var prevUninstallCommand = uninstallTheAppKey?.GetValue("UninstallString");
            uninstallTheAppKey?.SetValue("UninstallString.0", prevUninstallCommand);
            uninstallTheAppKey?.SetValue("UninstallString", $"\"{Environment.GetCommandLineArgs().First()}\" uninstall");
        }

        /// <summary>
        /// Execute custom uninstall process.
        /// </summary>
        private static void Uninstall()
        {
            // Retrieve ClickOnce original uninstall command.
            var uninstallTheAppKey = FindUninstallTheAppRegKey();
            var uninstallCommand = (uninstallTheAppKey?.GetValue("UninstallString.0")?.ToString())?.Split(' ');
            if (uninstallCommand == null) return;

            // Execute ClickOnce original uninstaller and wait for it's exit.
            var uninstallProcess = Process.Start(
                fileName: uninstallCommand.First(),
                arguments: string.Join(" ", uninstallCommand.Skip(1).ToArray()));

            uninstallProcess.WaitForExit();

            // If canceld uninstall or recover only, then nothing to do.
            if (FindUninstallTheAppRegKey() != null) return;

            // Delete "Open W3C log with Excel" verb.
            var HKCR = Registry.ClassesRoot;
            var txtShellKey = HKCR.OpenSubKey(@"txtfile\shell", writable: true);
            if (txtShellKey.GetSubKeyNames().Contains(verb))
            {
                HKCR.DeleteSubKey(verbPath + @"\command");
                HKCR.DeleteSubKey(verbPath);
            }

            // Restore default verb of "txtfile".
            var defaultVerb = txtShellKey.GetValue("").ToString();
            if (defaultVerb == verb)
            {
                txtShellKey.SetValue("", "open");
            }
        }

        /// <summary>
        /// Find registry key of uninstall this app.
        /// </summary>
        private static RegistryKey FindUninstallTheAppRegKey(bool writable = false)
        {
            var HKCU = Registry.CurrentUser;
            var uninstallRootKey = HKCU.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
            var uninstallTheAppKey = uninstallRootKey?
                .GetSubKeyNames()
                .Select(subKeyName => uninstallRootKey.OpenSubKey(subKeyName, writable))
                .FirstOrDefault(subKey => subKey.GetValue("DisplayName").ToString() == "Open W3C Extended Log with Excel");
            return uninstallTheAppKey;
        }
    }
}
