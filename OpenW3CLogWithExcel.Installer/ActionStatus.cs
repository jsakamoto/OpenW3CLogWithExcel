using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenW3CLogWithExcel.Installer
{
    [Flags]
    public enum ActionStatus
    {
        NotUninstalled  = 0,
        InstallSuccess = 1,
        AlreayInstalled = 2,
        VerbCommandNotFound = 4,
        CouldNotFoundUninstallKey = 8,
        OriginalUninstallCommandNotFound = 16,
        UninstallSuccess = 32
    }
}
