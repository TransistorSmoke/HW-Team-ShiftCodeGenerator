namespace HWAttendanceGrabber
{
    partial class CheckAttendanceWin
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
            this.btnCreateCSV = new System.Windows.Forms.Button();
            this.btnCheckAttFiles = new System.Windows.Forms.Button();
            this.labelCheckFile = new System.Windows.Forms.Label();
            this.labelCreateCSV = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // btnCreateCSV
            // 
            this.btnCreateCSV.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateCSV.Location = new System.Drawing.Point(12, 58);
            this.btnCreateCSV.Name = "btnCreateCSV";
            this.btnCreateCSV.Size = new System.Drawing.Size(117, 40);
            this.btnCreateCSV.TabIndex = 0;
            this.btnCreateCSV.TabStop = false;
            this.btnCreateCSV.Text = "CREATE CSV FILE";
            this.btnCreateCSV.UseVisualStyleBackColor = true;
            this.btnCreateCSV.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnCheckAttFiles
            // 
            this.btnCheckAttFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCheckAttFiles.Location = new System.Drawing.Point(12, 12);
            this.btnCheckAttFiles.Name = "btnCheckAttFiles";
            this.btnCheckAttFiles.Size = new System.Drawing.Size(117, 40);
            this.btnCheckAttFiles.TabIndex = 4;
            this.btnCheckAttFiles.TabStop = false;
            this.btnCheckAttFiles.Text = "CHECK FILES";
            this.btnCheckAttFiles.UseVisualStyleBackColor = true;
            this.btnCheckAttFiles.Click += new System.EventHandler(this.button2_Click);
            // 
            // labelCheckFile
            // 
            this.labelCheckFile.AutoSize = true;
            this.labelCheckFile.Location = new System.Drawing.Point(135, 19);
            this.labelCheckFile.Name = "labelCheckFile";
            this.labelCheckFile.Size = new System.Drawing.Size(247, 26);
            this.labelCheckFile.TabIndex = 5;
            this.labelCheckFile.Text = "Checks for attendance data files from the directory \r\ncurrently existing in the d" +
                "irectory";
            this.labelCheckFile.Click += new System.EventHandler(this.labelCheckFile_Click);
            // 
            // labelCreateCSV
            // 
            this.labelCreateCSV.AutoSize = true;
            this.labelCreateCSV.Location = new System.Drawing.Point(135, 65);
            this.labelCreateCSV.Name = "labelCreateCSV";
            this.labelCreateCSV.Size = new System.Drawing.Size(238, 26);
            this.labelCreateCSV.TabIndex = 6;
            this.labelCreateCSV.Text = "Generates the appropriate CSV file based on the \r\nattendance files present in the" +
                " directory";
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Location = new System.Drawing.Point(159, 113);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(76, 30);
            this.btnClose.TabIndex = 7;
            this.btnClose.TabStop = false;
            this.btnClose.Text = "CLOSE";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            // 
            // grabAttendanceWin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(381, 153);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.labelCreateCSV);
            this.Controls.Add(this.labelCheckFile);
            this.Controls.Add(this.btnCheckAttFiles);
            this.Controls.Add(this.btnCreateCSV);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "grabAttendanceWin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ATTENDANCE DATA CHECK";
            this.Load += new System.EventHandler(this.grabAttendanceWin_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCreateCSV;
        private System.Windows.Forms.Button btnCheckAttFiles;
        private System.Windows.Forms.Label labelCheckFile;
        private System.Windows.Forms.Label labelCreateCSV;
        private System.Windows.Forms.Button btnClose;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;

    }
}