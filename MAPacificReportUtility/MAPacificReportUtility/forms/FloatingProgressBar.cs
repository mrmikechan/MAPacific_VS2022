using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

///FloatingProgressBar class is a basic form used to create a dialog box to display the progress of
///a process intensive action.
namespace MAPacificReportUtility.forms
{
    public partial class FloatingProgressBar : Form
    {
        public int PercentageComplete { get; set; } = 0;
        public FloatingProgressBar()
        {
            InitializeComponent();
        }

        public FloatingProgressBar(string intxt, int max)
        {
            InitializeComponent();
            label1.Text = intxt;
            progressBar1.Maximum = max;
            progressBar1.Minimum = 0;
        }

        public void UpdateProgressBar(int progress)
        {
            if (progressBar1.InvokeRequired)
            {
                progressBar1.BeginInvoke(new Action(() => progressBar1.Value = progress));
            }
            else
                progressBar1.Value = progress;
        }
    }
}
