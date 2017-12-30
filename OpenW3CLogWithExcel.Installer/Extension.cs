using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace OpenW3CLogWithExcel.Installer
{
    internal static class Extension
    {
        public static bool HasFlags(this ActionStatus value, ActionStatus test)
        {
            return ((value & test) == test);
        }

        public static void DeleteSubKeyTreeAnyway(this RegistryKey registryKey, string subkey)
        {
            try { registryKey.DeleteSubKeyTree(subkey); }
            catch (ArgumentException) { }
        }
    }
}
