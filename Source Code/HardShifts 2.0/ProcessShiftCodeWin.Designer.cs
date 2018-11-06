namespace HWAttendanceGrabber
{
    partial class ProcessShiftCodeWin
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
            this.btnGetData = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.yearTextBox = new System.Windows.Forms.TextBox();
            this.ddSection = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.ddTo = new System.Windows.Forms.ComboBox();
            this.labelTo = new System.Windows.Forms.Label();
            this.ddFrom = new System.Windows.Forms.ComboBox();
            this.labelFrom = new System.Windows.Forms.Label();
            this.labelSpecify = new System.Windows.Forms.Label();
            this.radioButtonNo = new System.Windows.Forms.RadioButton();
            this.radioButtonYes = new System.Windows.Forms.RadioButton();
            this.label5 = new System.Windows.Forms.Label();
            this.ddShiftPeriod = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.titleLabel = new System.Windows.Forms.Label();
            this.SelectFileBtn = new System.Windows.Forms.Button();
            this.fileNameBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnClose = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGetData
            // 
            this.btnGetData.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGetData.Location = new System.Drawing.Point(26, 400);
            this.btnGetData.Name = "btnGetData";
            this.btnGetData.Size = new System.Drawing.Size(186, 37);
            this.btnGetData.TabIndex = 0;
            this.btnGetData.Text = "GENERATE SHIFT CODES";
            this.btnGetData.UseVisualStyleBackColor = true;
            this.btnGetData.Click += new System.EventHandler(this.btnGetData_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.yearTextBox);
            this.panel1.Controls.Add(this.ddSection);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.ddTo);
            this.panel1.Controls.Add(this.labelTo);
            this.panel1.Controls.Add(this.ddFrom);
            this.panel1.Controls.Add(this.labelFrom);
            this.panel1.Controls.Add(this.labelSpecify);
            this.panel1.Controls.Add(this.radioButtonNo);
            this.panel1.Controls.Add(this.radioButtonYes);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.ddShiftPeriod);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Location = new System.Drawing.Point(12, 180);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(410, 205);
            this.panel1.TabIndex = 9;
            // 
            // yearTextBox
            // 
            this.yearTextBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.yearTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yearTextBox.Location = new System.Drawing.Point(218, 7);
            this.yearTextBox.Name = "yearTextBox";
            this.yearTextBox.ReadOnly = true;
            this.yearTextBox.Size = new System.Drawing.Size(171, 20);
            this.yearTextBox.TabIndex = 14;
            this.yearTextBox.TabStop = false;
            this.yearTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // ddSection
            // 
            this.ddSection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ddSection.FormattingEnabled = true;
            this.ddSection.Items.AddRange(new object[] {
            "[ALL SECTIONS]",
            "HARDWARE DEVELOPMENT",
            "SOFTWARE DEVELOPMENT",
            "SOFTWARE VALIDATION",
            "ELECTROMECHANICAL",
            "LAB, PM AND PQ"});
            this.ddSection.Location = new System.Drawing.Point(219, 68);
            this.ddSection.Name = "ddSection";
            this.ddSection.Size = new System.Drawing.Size(170, 21);
            this.ddSection.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(10, 68);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(80, 17);
            this.label4.TabIndex = 12;
            this.label4.Text = "SECTION:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(10, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 17);
            this.label3.TabIndex = 10;
            this.label3.Text = "YEAR:";
            // 
            // ddTo
            // 
            this.ddTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ddTo.FormattingEnabled = true;
            this.ddTo.Location = new System.Drawing.Point(242, 157);
            this.ddTo.Name = "ddTo";
            this.ddTo.Size = new System.Drawing.Size(114, 21);
            this.ddTo.TabIndex = 9;
            // 
            // labelTo
            // 
            this.labelTo.AutoSize = true;
            this.labelTo.Location = new System.Drawing.Point(218, 160);
            this.labelTo.Name = "labelTo";
            this.labelTo.Size = new System.Drawing.Size(28, 13);
            this.labelTo.TabIndex = 8;
            this.labelTo.Text = "TO: ";
            // 
            // ddFrom
            // 
            this.ddFrom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ddFrom.FormattingEnabled = true;
            this.ddFrom.Location = new System.Drawing.Point(71, 157);
            this.ddFrom.Name = "ddFrom";
            this.ddFrom.Size = new System.Drawing.Size(114, 21);
            this.ddFrom.TabIndex = 7;
            // 
            // labelFrom
            // 
            this.labelFrom.AutoSize = true;
            this.labelFrom.Location = new System.Drawing.Point(30, 160);
            this.labelFrom.Name = "labelFrom";
            this.labelFrom.Size = new System.Drawing.Size(41, 13);
            this.labelFrom.TabIndex = 6;
            this.labelFrom.Text = "FROM:";
            // 
            // labelSpecify
            // 
            this.labelSpecify.AutoSize = true;
            this.labelSpecify.Location = new System.Drawing.Point(10, 128);
            this.labelSpecify.Name = "labelSpecify";
            this.labelSpecify.Size = new System.Drawing.Size(128, 13);
            this.labelSpecify.TabIndex = 5;
            this.labelSpecify.Text = "If yes, specify date range:";
            // 
            // radioButtonNo
            // 
            this.radioButtonNo.AutoSize = true;
            this.radioButtonNo.Checked = true;
            this.radioButtonNo.Location = new System.Drawing.Point(317, 109);
            this.radioButtonNo.Name = "radioButtonNo";
            this.radioButtonNo.Size = new System.Drawing.Size(39, 17);
            this.radioButtonNo.TabIndex = 3;
            this.radioButtonNo.TabStop = true;
            this.radioButtonNo.Text = "No";
            this.radioButtonNo.UseVisualStyleBackColor = true;
            this.radioButtonNo.CheckedChanged += new System.EventHandler(this.radioButtonNo_CheckedChanged);
            // 
            // radioButtonYes
            // 
            this.radioButtonYes.AutoSize = true;
            this.radioButtonYes.Location = new System.Drawing.Point(268, 109);
            this.radioButtonYes.Name = "radioButtonYes";
            this.radioButtonYes.Size = new System.Drawing.Size(43, 17);
            this.radioButtonYes.TabIndex = 2;
            this.radioButtonYes.TabStop = true;
            this.radioButtonYes.Text = "Yes";
            this.radioButtonYes.UseVisualStyleBackColor = true;
            this.radioButtonYes.CheckedChanged += new System.EventHandler(this.radioButtonYes_CheckedChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 109);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(234, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Are there 8-hour working days within the period?";
            // 
            // ddShiftPeriod
            // 
            this.ddShiftPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ddShiftPeriod.FormattingEnabled = true;
            this.ddShiftPeriod.Location = new System.Drawing.Point(219, 39);
            this.ddShiftPeriod.Name = "ddShiftPeriod";
            this.ddShiftPeriod.Size = new System.Drawing.Size(170, 21);
            this.ddShiftPeriod.TabIndex = 1;
            this.ddShiftPeriod.SelectedIndexChanged += new System.EventHandler(this.ddShiftPeriod_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(10, 39);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(120, 17);
            this.label6.TabIndex = 0;
            this.label6.Text = "SHIFT PERIOD:";
            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLabel.Location = new System.Drawing.Point(64, 9);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(339, 20);
            this.titleLabel.TabIndex = 8;
            this.titleLabel.Text = "AUTOMATIC SHIFT CODE GENERATOR";
            // 
            // SelectFileBtn
            // 
            this.SelectFileBtn.Font = new System.Drawing.Font("Arial Narrow", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SelectFileBtn.Location = new System.Drawing.Point(152, 86);
            this.SelectFileBtn.Name = "SelectFileBtn";
            this.SelectFileBtn.Size = new System.Drawing.Size(108, 32);
            this.SelectFileBtn.TabIndex = 10;
            this.SelectFileBtn.Text = "SELECT FILE";
            this.SelectFileBtn.UseVisualStyleBackColor = true;
            this.SelectFileBtn.Click += new System.EventHandler(this.SelectFileBtn_Click);
            // 
            // fileNameBox
            // 
            this.fileNameBox.Location = new System.Drawing.Point(6, 60);
            this.fileNameBox.Name = "fileNameBox";
            this.fileNameBox.Size = new System.Drawing.Size(387, 20);
            this.fileNameBox.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(155, 17);
            this.label1.TabIndex = 12;
            this.label1.Text = "ATTENDANCE FILE NAME:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(159, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Please select attendance sheet:";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.SelectFileBtn);
            this.panel2.Controls.Add(this.fileNameBox);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(12, 41);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(410, 129);
            this.panel2.TabIndex = 14;
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Location = new System.Drawing.Point(238, 400);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(168, 37);
            this.btnClose.TabIndex = 15;
            this.btnClose.Text = "CLOSE";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // ProcessShiftCodeWin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(435, 456);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.titleLabel);
            this.Controls.Add(this.btnGetData);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ProcessShiftCodeWin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generate Shift Code";
            this.Load += new System.EventHandler(this.ProcessShiftCodeWin_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGetData;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox ddSection;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox ddTo;
        private System.Windows.Forms.Label labelTo;
        private System.Windows.Forms.ComboBox ddFrom;
        private System.Windows.Forms.Label labelFrom;
        private System.Windows.Forms.Label labelSpecify;
        private System.Windows.Forms.RadioButton radioButtonNo;
        private System.Windows.Forms.RadioButton radioButtonYes;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox ddShiftPeriod;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Button SelectFileBtn;
        private System.Windows.Forms.TextBox fileNameBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox yearTextBox;
        private System.Windows.Forms.Button btnClose;
    }
}