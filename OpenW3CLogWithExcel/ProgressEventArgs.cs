using System;

namespace OpenW3CLogWithExcel
{
    public class ProgressEventArgs : EventArgs
    {
        public int Progress { get; }

        public ProgressEventArgs(int progress)
        {
            this.Progress = progress;
        }
    }
}