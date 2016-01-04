using System;
using System.Collections.Generic;
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
            var commandLineArgs = ApplicationDeployment.IsNetworkDeployed ? AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData ?? new string[0] : args;

            // Validate command line arguments.

            if (commandLineArgs.Any() == false)
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var installed = Registry.ClassesRoot.OpenSubKey(@"txtfile\shell\openw3clogwithexcel") != null;
                    if (installed) return;
                    var installerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OpenW3CLogWithExcel.Installer.exe");
                    Process.Start(installerPath, "install");
                }
                return;
            }

            var path = commandLineArgs.First();
            if (File.Exists(path) == false) return;

            new W3CLogOpener().Open(path);
        }
    }
}
