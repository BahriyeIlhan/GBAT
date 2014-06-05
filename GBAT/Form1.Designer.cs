namespace GBAT
{
    partial class Form1
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
            this.IFCFileLabel = new System.Windows.Forms.Label();
            this.IFCFileUrl = new System.Windows.Forms.TextBox();
            this.BrowseIFCFile = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.MaterialList = new System.Windows.Forms.CheckedListBox();
            this.ProcessIFCFile = new System.Windows.Forms.Button();
            this.Results = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.CancelProcess = new System.Windows.Forms.Button();
            this.IFAProgress = new System.Windows.Forms.ProgressBar();
            this.SelectAll = new System.Windows.Forms.Button();
            this.DeselectAll = new System.Windows.Forms.Button();
            this.Report = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // IFCFileLabel
            // 
            this.IFCFileLabel.AutoSize = true;
            this.IFCFileLabel.Location = new System.Drawing.Point(15, 26);
            this.IFCFileLabel.Name = "IFCFileLabel";
            this.IFCFileLabel.Size = new System.Drawing.Size(42, 13);
            this.IFCFileLabel.TabIndex = 4;
            this.IFCFileLabel.Text = "IFC File";
            // 
            // IFCFileUrl
            // 
            this.IFCFileUrl.Location = new System.Drawing.Point(60, 22);
            this.IFCFileUrl.Name = "IFCFileUrl";
            this.IFCFileUrl.Size = new System.Drawing.Size(350, 20);
            this.IFCFileUrl.TabIndex = 5;
            // 
            // BrowseIFCFile
            // 
            this.BrowseIFCFile.Location = new System.Drawing.Point(416, 21);
            this.BrowseIFCFile.Name = "BrowseIFCFile";
            this.BrowseIFCFile.Size = new System.Drawing.Size(75, 22);
            this.BrowseIFCFile.TabIndex = 6;
            this.BrowseIFCFile.Text = "Browse";
            this.BrowseIFCFile.UseVisualStyleBackColor = true;
            this.BrowseIFCFile.Click += new System.EventHandler(this.BrowseIfcFileClicked);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.MaterialList);
            this.groupBox1.Location = new System.Drawing.Point(16, 68);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(562, 102);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Materials";
            // 
            // MaterialList
            // 
            this.MaterialList.BackColor = System.Drawing.SystemColors.Menu;
            this.MaterialList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.MaterialList.CheckOnClick = true;
            this.MaterialList.ColumnWidth = 270;
            this.MaterialList.FormattingEnabled = true;
            this.MaterialList.Location = new System.Drawing.Point(7, 20);
            this.MaterialList.MultiColumn = true;
            this.MaterialList.Name = "MaterialList";
            this.MaterialList.Size = new System.Drawing.Size(548, 75);
            this.MaterialList.TabIndex = 0;
            // 
            // ProcessIFCFile
            // 
            this.ProcessIFCFile.Location = new System.Drawing.Point(16, 216);
            this.ProcessIFCFile.Name = "ProcessIFCFile";
            this.ProcessIFCFile.Size = new System.Drawing.Size(96, 23);
            this.ProcessIFCFile.TabIndex = 8;
            this.ProcessIFCFile.Text = "Process";
            this.ProcessIFCFile.UseVisualStyleBackColor = true;
            this.ProcessIFCFile.Click += new System.EventHandler(this.ProcessIfcFileClicked);
            // 
            // Results
            // 
            this.Results.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Results.Location = new System.Drawing.Point(6, 18);
            this.Results.Multiline = true;
            this.Results.Name = "Results";
            this.Results.ReadOnly = true;
            this.Results.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.Results.Size = new System.Drawing.Size(296, 108);
            this.Results.TabIndex = 9;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.Results);
            this.groupBox2.Location = new System.Drawing.Point(270, 216);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(308, 141);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Result";
            // 
            // CancelProcess
            // 
            this.CancelProcess.Enabled = false;
            this.CancelProcess.Location = new System.Drawing.Point(118, 216);
            this.CancelProcess.Name = "CancelProcess";
            this.CancelProcess.Size = new System.Drawing.Size(100, 23);
            this.CancelProcess.TabIndex = 11;
            this.CancelProcess.Text = "Cancel";
            this.CancelProcess.UseVisualStyleBackColor = true;
            this.CancelProcess.Click += new System.EventHandler(this.CancelProcessClicked);
            // 
            // IFAProgress
            // 
            this.IFAProgress.Location = new System.Drawing.Point(16, 246);
            this.IFAProgress.Name = "IFAProgress";
            this.IFAProgress.Size = new System.Drawing.Size(202, 12);
            this.IFAProgress.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.IFAProgress.TabIndex = 12;
            this.IFAProgress.Visible = false;
            // 
            // SelectAll
            // 
            this.SelectAll.Location = new System.Drawing.Point(422, 176);
            this.SelectAll.Name = "SelectAll";
            this.SelectAll.Size = new System.Drawing.Size(75, 23);
            this.SelectAll.TabIndex = 13;
            this.SelectAll.Text = "Select All";
            this.SelectAll.UseVisualStyleBackColor = true;
            this.SelectAll.Click += new System.EventHandler(this.SelectAllClicked);
            // 
            // DeselectAll
            // 
            this.DeselectAll.Location = new System.Drawing.Point(503, 176);
            this.DeselectAll.Name = "DeselectAll";
            this.DeselectAll.Size = new System.Drawing.Size(75, 23);
            this.DeselectAll.TabIndex = 14;
            this.DeselectAll.Text = "Deselect All";
            this.DeselectAll.UseVisualStyleBackColor = true;
            this.DeselectAll.Click += new System.EventHandler(this.DeselectAllClicked);
            // 
            // Report
            // 
            this.Report.Location = new System.Drawing.Point(18, 265);
            this.Report.Name = "Report";
            this.Report.Size = new System.Drawing.Size(75, 23);
            this.Report.TabIndex = 15;
            this.Report.Text = "Save Report";
            this.Report.UseVisualStyleBackColor = true;
            this.Report.Visible = false;
            this.Report.Click += new System.EventHandler(this.ReportClicked);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(590, 369);
            this.Controls.Add(this.Report);
            this.Controls.Add(this.DeselectAll);
            this.Controls.Add(this.SelectAll);
            this.Controls.Add(this.IFAProgress);
            this.Controls.Add(this.CancelProcess);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.ProcessIFCFile);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BrowseIFCFile);
            this.Controls.Add(this.IFCFileUrl);
            this.Controls.Add(this.IFCFileLabel);
            this.Name = "Form1";
            this.Text = "GBAT";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label IFCFileLabel;
        private System.Windows.Forms.TextBox IFCFileUrl;
        private System.Windows.Forms.Button BrowseIFCFile;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckedListBox MaterialList;
        private System.Windows.Forms.Button ProcessIFCFile;
        private System.Windows.Forms.TextBox Results;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button CancelProcess;
        private System.Windows.Forms.ProgressBar IFAProgress;
        private System.Windows.Forms.Button SelectAll;
        private System.Windows.Forms.Button DeselectAll;
        private System.Windows.Forms.Button Report;
    }
}

