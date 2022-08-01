namespace MAPacificReportUtility.forms
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxDirectory = new System.Windows.Forms.TextBox();
            this.bindingSourceUserSettings = new System.Windows.Forms.BindingSource(this.components);
            this.buttonFolder = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonFolderBrInfo = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btnEmailDraft = new System.Windows.Forms.Button();
            this.txtBoxEmailDraft = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSourceUserSettings)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 5;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.textBoxDirectory, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.buttonFolder, 4, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.textBox1, 2, 2);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.buttonFolderBrInfo, 2, 4);
            this.tableLayoutPanel1.Controls.Add(this.buttonOK, 2, 6);
            this.tableLayoutPanel1.Controls.Add(this.buttonCancel, 3, 6);
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.btnEmailDraft, 4, 5);
            this.tableLayoutPanel1.Controls.Add(this.txtBoxEmailDraft, 2, 5);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(6, 6);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(6);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.Padding = new System.Windows.Forms.Padding(3);
            this.tableLayoutPanel1.RowCount = 7;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(439, 160);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Excel Output Directory";
            this.toolTip1.SetToolTip(this.label2, "Enter the directory where the Excel file will be saved.");
            // 
            // textBoxDirectory
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.textBoxDirectory, 2);
            this.textBoxDirectory.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSourceUserSettings, "ExcelDirectory", true));
            this.textBoxDirectory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxDirectory.Location = new System.Drawing.Point(138, 6);
            this.textBoxDirectory.Name = "textBoxDirectory";
            this.textBoxDirectory.Size = new System.Drawing.Size(217, 20);
            this.textBoxDirectory.TabIndex = 3;
            // 
            // bindingSourceUserSettings
            // 
            this.bindingSourceUserSettings.DataSource = typeof(MAPacificReportUtility.UserSettings);
            // 
            // buttonFolder
            // 
            this.buttonFolder.Location = new System.Drawing.Point(361, 6);
            this.buttonFolder.Name = "buttonFolder";
            this.buttonFolder.Size = new System.Drawing.Size(75, 23);
            this.buttonFolder.TabIndex = 6;
            this.buttonFolder.Text = "Folder...";
            this.buttonFolder.UseVisualStyleBackColor = true;
            this.buttonFolder.Click += new System.EventHandler(this.buttonFolder_Click);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Excel File Name Prefix";
            // 
            // textBox1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.textBox1, 2);
            this.textBox1.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSourceUserSettings, "ExcelFileNamePrefix", true));
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Location = new System.Drawing.Point(138, 40);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(217, 20);
            this.textBox1.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 76);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(121, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "BranchInfo.xml Location";
            // 
            // buttonFolderBrInfo
            // 
            this.buttonFolderBrInfo.Location = new System.Drawing.Point(138, 71);
            this.buttonFolderBrInfo.Name = "buttonFolderBrInfo";
            this.buttonFolderBrInfo.Size = new System.Drawing.Size(75, 23);
            this.buttonFolderBrInfo.TabIndex = 10;
            this.buttonFolderBrInfo.Text = "Folder...";
            this.buttonFolderBrInfo.UseVisualStyleBackColor = true;
            this.buttonFolderBrInfo.Click += new System.EventHandler(this.buttonFolderBrInfo_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.buttonOK.Location = new System.Drawing.Point(138, 130);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 22);
            this.buttonOK.TabIndex = 4;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(219, 130);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 22);
            this.buttonCancel.TabIndex = 5;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 105);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(103, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Email Draft Directory";
            this.label4.Visible = false;
            // 
            // btnEmailDraft
            // 
            this.btnEmailDraft.Location = new System.Drawing.Point(361, 100);
            this.btnEmailDraft.Name = "btnEmailDraft";
            this.btnEmailDraft.Size = new System.Drawing.Size(75, 23);
            this.btnEmailDraft.TabIndex = 12;
            this.btnEmailDraft.Text = "Folder...";
            this.btnEmailDraft.UseVisualStyleBackColor = true;
            this.btnEmailDraft.Visible = false;
            this.btnEmailDraft.Click += new System.EventHandler(this.buttonEmailFolder_Click);
            // 
            // txtBoxEmailDraft
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.txtBoxEmailDraft, 2);
            this.txtBoxEmailDraft.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSourceUserSettings, "DraftEmailPath", true));
            this.txtBoxEmailDraft.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtBoxEmailDraft.Location = new System.Drawing.Point(138, 100);
            this.txtBoxEmailDraft.Name = "txtBoxEmailDraft";
            this.txtBoxEmailDraft.Size = new System.Drawing.Size(217, 20);
            this.txtBoxEmailDraft.TabIndex = 13;
            this.txtBoxEmailDraft.Visible = false;
            // 
            // toolTip1
            // 
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(451, 172);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "SettingsForm";
            this.Padding = new System.Windows.Forms.Padding(6);
            this.Text = "Application Settings";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSourceUserSettings)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxDirectory;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.BindingSource bindingSourceUserSettings;
        private System.Windows.Forms.Button buttonFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonFolderBrInfo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnEmailDraft;
        private System.Windows.Forms.TextBox txtBoxEmailDraft;
    }
}