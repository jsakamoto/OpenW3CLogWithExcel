using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
            this.W3CLogOpener.Exit += W3CLogOpener_Exit;
            this.W3CLogOpener.Converted += W3CLogOpener_Converted;
        }

        private void W3CLogOpener_Converted(object sender, EventArgs e)
        {
            this.Visible = false;
        }

        private void W3CLogOpener_Exit(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            new Thread(() => this.W3CLogOpener.Open(LogPath))
                .Start();
        }
    }
}
