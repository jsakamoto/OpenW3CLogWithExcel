using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace OpenW3CLogWithExcel
{
    public partial class MainForm : Form
    {
        private string LogPath { get; set; }

        private W3CLogOpener W3CLogOpener { get; } = new W3CLogOpener();

        public MainForm()
        {
            InitializeComponent();
        }

        public MainForm(string logPath)
        {
            InitializeComponent();
            LogPath = logPath;
            this.W3CLogOpener.Progress += W3CLogOpener_Progress;
            this.W3CLogOpener.Exit += W3CLogOpener_Exit;
            this.W3CLogOpener.Converted += W3CLogOpener_Converted;
        }

        private void Invoke(Action action)
        {
            this.Invoke((MethodInvoker)(() => action()));
        }

        private void W3CLogOpener_Progress(object sender, ProgressEventArgs e)
        {
            Invoke(() => this.progressBar1.Value = e.Progress);
        }

        private void W3CLogOpener_Converted(object sender, EventArgs e)
        {
            Invoke(() => this.Visible = false);
        }

        private void W3CLogOpener_Exit(object sender, EventArgs e)
        {
            Invoke(() => this.Close());
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            new Thread(() => this.W3CLogOpener.Open(LogPath))
                .Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.timer1.Stop();
            this.Opacity = 1.0;
        }
    }
}
