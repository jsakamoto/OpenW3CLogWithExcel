using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenW3CLogWithExcel.Installer
{
    internal static class Extension
    {
        public static bool HasFlags(this ActionStatus value, ActionStatus test) 
        {
            return ((value & test) == test);
        }
    }
}
