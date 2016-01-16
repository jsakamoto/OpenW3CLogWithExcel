using System;
using System.Diagnostics;

namespace OpenW3CLogWithExcel
{
    public interface IShell
    {
        Process Open(string path);
    }
}
