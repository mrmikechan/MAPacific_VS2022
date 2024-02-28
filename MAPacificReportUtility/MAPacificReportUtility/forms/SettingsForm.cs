using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


///
/// SettingsForm Class
/// 
/// Form object class that contains User Customizable settings that is databound to UserSettings.
/// 
namespace MAPacificReportUtility.forms
{
    public partial class SettingsForm : Form
    {
        /// <summary>
        /// SettingsForm Constructor initialize the Form object and set the binding source to UserSettings
        /// </summary>
        public SettingsForm()
        {
            InitializeComponent();
            UserSettings.Save();
            bindingSourceUserSettings.DataSource = UserSettings.Current;

        }

        /// <summary>
        /// buttonOK_Click event saves the customized data.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonOK_Click(object sender, EventArgs e)
        {
            UserSettings.Save();
            DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// buttonCancel_Click event listener closes the dialog box without saving the customized data.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonCancel_Click(object sender, EventArgs e)
        {

        }

        private void buttonEmailFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = UserSettings.Current.DraftEmailPath;
            DialogResult result = folderBrowserDialog1.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                txtBoxEmailDraft.Text = folderBrowserDialog1.SelectedPath;
                UserSettings.Current.DraftEmailPath = folderBrowserDialog1.SelectedPath;
            }
        }

        /// <summary>
        /// buttonFolder_Click event listener launches the folder browser dialog box so users can
        /// select the folder path in a GUI container instead of manually typing the value.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonFolder_Click(object sender, EventArgs e)
        {

            folderBrowserDialog1.SelectedPath = UserSettings.Current.ExcelDirectory;
            DialogResult result = folderBrowserDialog1.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                textBoxDirectory.Text = folderBrowserDialog1.SelectedPath;
                UserSettings.Current.ExcelDirectory = folderBrowserDialog1.SelectedPath;
            }
        }

        /// <summary>
        /// Open up windows explorer and display the folder location that the BranchINfo.xml is located.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonFolderBrInfo_Click(object sender, EventArgs e)
        {
            //check roaming profile first...
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string path = appData + System.IO.Path.DirectorySeparatorChar + "MAPReportUtility";
            bool pathExist = false;
            if(System.IO.File.Exists(path + System.IO.Path.DirectorySeparatorChar + "BranchInfo.xml"))
            {
                pathExist = true;
            }

            if (!pathExist)
            {
                appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                path = appData + System.IO.Path.DirectorySeparatorChar + "MAPReportUtility";
                if (System.IO.File.Exists(path + System.IO.Path.DirectorySeparatorChar + "BranchInfo.xml"))
                {
                    pathExist = true;
                }
                else //Generate a MessageBox to notify the user that thise file does not exist
                {
                    MessageBox.Show($"BranchInfo.xml does not exist in the following location: {path}{System.IO.Path.DirectorySeparatorChar}","Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }

            if (pathExist)
            {
                try
                {
                    System.Diagnostics.Process prc = new System.Diagnostics.Process();
                    string windir = Environment.GetEnvironmentVariable("WINDIR");
                    prc.StartInfo.FileName = windir + System.IO.Path.DirectorySeparatorChar + @"explorer.exe";
                    prc.StartInfo.Arguments = path;
                    prc.Start();
                }
                catch (Exception ex1) 
                {
                    string error = ex1.Message;
                }
            }
        }
    }
}
