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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var status = ActionStatus.NotUninstalled;
            switch (args.FirstOrDefault())
            {
                case "install":
                    status = Install(args.LastOrDefault() ?? "1.0.0.0");
                    break;
                case "uninstall":
                    status = Uninstall();
                    break;
                default:
                    break;
            }

            ShowResultMessageBox(status);
        }

        private const string verb = "openw3clogwithexcel";

        private static string verbPath => $@"txtfile\shell\{verb}";

        /// <summary>
        /// Execute install prcoess.
        /// </summary>
        private static ActionStatus Install(string appVer)
        {
            // Check installed or not
            var regKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\OpenW3CLogWithExcel");
            var installedVer = regKey?.GetValue("ver", "").ToString();
            if (installedVer == appVer) return ActionStatus.AlreayInstalled;

            // Add "Open W3C log with Excel" verb by cloning command which ClickOnce created.
            var HKCR = Registry.ClassesRoot;
            var commandKey = HKCR.OpenSubKey(@"0ea92b0253b34c489be1c97\shell\open\command");
            var commandValue = commandKey?.GetValue("")?.ToString();
            if (commandValue != null)
            {
                var shellKey = HKCR.CreateSubKey(verbPath);
                shellKey.SetValue("", "Open W3C log with Excel");

                var openLogCommanKey = shellKey.CreateSubKey("command");
                openLogCommanKey.SetValue("", commandValue);

                // Set "Open W3C log with Excel" to default verb.
                HKCR.OpenSubKey(@"txtfile\shell", writable: true).SetValue("", verb);
            }

            // Override uninstall command, from ClickOnce original uninstaller to this custom uninstall process.
            var uninstallTheAppKey = FindUninstallTheAppRegKey(writable: true);
            var prevUninstallCommand = uninstallTheAppKey?.GetValue("UninstallString")?.ToString();
            if (prevUninstallCommand.EndsWith(".exe\" uninstall") == false)
            {
                uninstallTheAppKey?.SetValue("UninstallString.0", prevUninstallCommand);
                uninstallTheAppKey?.SetValue("UninstallString", $"\"{Environment.GetCommandLineArgs().First()}\" uninstall");
            }

            // Write installed version information to HKCU Registory.
            Registry.CurrentUser.CreateSubKey(@"SOFTWARE\OpenW3CLogWithExcel")
                .SetValue("ver", appVer);

            var status = ActionStatus.NotUninstalled;
            if (commandValue == null) status |= ActionStatus.VerbCommandNotFound;
            if (uninstallTheAppKey == null) status |= ActionStatus.CouldNotFoundUninstallKey;
            if (status == ActionStatus.NotUninstalled) status = ActionStatus.InstallSuccess;
            return status;
        }

        /// <summary>
        /// Execute custom uninstall process.
        /// </summary>
        private static ActionStatus Uninstall()
        {
            var HKCR = Registry.ClassesRoot;

            // Retrieve ClickOnce original uninstall command.
            var uninstallTheAppKey = FindUninstallTheAppRegKey();
            var uninstallCommand = (uninstallTheAppKey?.GetValue("UninstallString.0")?.ToString())?.Split(' ');
            if (uninstallCommand == null) return ActionStatus.OriginalUninstallCommandNotFound;

            // Execute ClickOnce original uninstaller and wait for it's exit.
            var uninstallProcess = Process.Start(
                fileName: uninstallCommand.First(),
                arguments: string.Join(" ", uninstallCommand.Skip(1).ToArray()));

            uninstallProcess.WaitForExit();

            // If canceld uninstall or recover only...
            if (FindUninstallTheAppRegKey() != null)
            {
                // re-launch OpenW3CLogWithExcel command to fix/recover registry settings.
                var commandParts = HKCR.OpenSubKey(@"0ea92b0253b34c489be1c97\shell\open\command")
                    .GetValue("")
                    .ToString()
                    .Split(' ')
                    .Where(s => s != "%1")
                    .ToArray();
                Process.Start(
                    fileName: commandParts.First(),
                    arguments: string.Join(" ", commandParts.Skip(1).ToArray()));

                return ActionStatus.NotUninstalled;
            }

            // Delete "Open W3C log with Excel" verb.
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

            // Remove installed version information to HKCU Registory.
            Registry.CurrentUser.DeleteSubKeyTree(@"SOFTWARE\OpenW3CLogWithExcel");

            return ActionStatus.UninstallSuccess;
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

        private static void ShowResultMessageBox(ActionStatus status)
        {
            var message = "";
            var icon = MessageBoxIcon.Information;

            if (status.HasFlags(ActionStatus.AlreayInstalled))
            {
                message = "The application seems to be already installed.";
            }
            if (status.HasFlags(ActionStatus.InstallSuccess))
            {
                message = "The application was installed successfully.";
            }
            if (status.HasFlags(ActionStatus.VerbCommandNotFound))
            {
                icon = MessageBoxIcon.Error;
                message += @"""HKEY_CLASSES_ROOT\0ea92b0253b34c489be1c97\shell\open\command"" registry key does not found." + "\n\n";
            }
            if (status.HasFlags(ActionStatus.CouldNotFoundUninstallKey))
            {
                icon = MessageBoxIcon.Error;
                message += @"Uninstall registry key does not found. (The subkey which is under ""HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"" and contains ""DisplayName""=""Open W3C Extended Log with Excel"" value)" + "\n\n";
            }
            if (status.HasFlags(ActionStatus.OriginalUninstallCommandNotFound))
            {
                icon = MessageBoxIcon.Error;
                message += @"Uninstall command does not found. (The command should be writen at ""UninstallString.0"" value in uninstall registry key)" + "\n\n";
            }

            if (message != "")
            {
                MessageBox.Show(message.TrimEnd('\n'), "Open W3C Extended Log with Excel", MessageBoxButtons.OK, icon);
            }
        }
    }
}
