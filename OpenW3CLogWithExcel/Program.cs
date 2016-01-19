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
                var thisVer = typeof(Program).Assembly.GetName().Version;
                var thisVerStr = $"{thisVer.Major}.{thisVer.Minor}.{thisVer.Build}.{thisVer.Revision}";

                var installedVer = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\OpenW3CLogWithExcel")?
                    .GetValue("ver", "")
                    .ToString();

                if (installedVer != thisVerStr)
                {
                    var installerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OpenW3CLogWithExcel.Installer.exe");
                    Process.Start(installerPath, "install " + thisVerStr).WaitForExit();
                }
            }

            if (commandLineArgs.Any() == false) return;

            var path = new Uri(commandLineArgs.First()).LocalPath;
            if (File.Exists(path) == false) return;

            Application.Run(new MainForm(path));
        }
    }
}
