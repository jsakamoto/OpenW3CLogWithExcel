using System;
using System.Deployment.Application;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OpenW3CLogWithExcel
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

            var commandLineArgs = ApplicationDeployment.IsNetworkDeployed ? AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData ?? new string[0] : args;

            // Check install condition.
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                InstallIfNecessary();
            }

            if (commandLineArgs.Any() == false) return;

            var path = new Uri(commandLineArgs.First()).LocalPath;
            if (File.Exists(path) == false) return;

            Application.Run(new MainForm(path));
        }

        private static void InstallIfNecessary()
        {
            var cmdExpected = Registry.GetValue(@"HKEY_CLASSES_ROOT\0ea92b0253b34c489be1c97\shell\open\command", "", null) as string;
            var cmdActual = Registry.GetValue(@"HKEY_CLASSES_ROOT\txtfile\shell\openw3clogwithexcel\command", "", null) as string;
            if (cmdExpected != cmdActual)
            {
                var installerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OpenW3CLogWithExcel.Installer.exe");
                Process.Start(installerPath, "install").WaitForExit();
            }
        }
    }
}
