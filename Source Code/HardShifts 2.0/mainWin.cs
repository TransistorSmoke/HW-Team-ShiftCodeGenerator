using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace HWAttendanceGrabber
{
    public partial class GrabForm : Form
    {
        public GrabForm()
        {
            InitializeComponent();                                          
        }

        private void closeApplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void GrabForm_Load(object sender, EventArgs e)
        {

        }


        private void generateShiftCodesToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            CheckAttendanceWin newAttChildWin = new CheckAttendanceWin();
            // Set the Parent Form of the Child window.
            newAttChildWin.MdiParent = this;
            // Display the new form.
            newAttChildWin.Show(); 
        }

        private void generateShiftCodesToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ProcessShiftCodeWin newProcShiftCodeWin = new ProcessShiftCodeWin();          
            // Set the Parent Form of the Child window.
            newProcShiftCodeWin.MdiParent = this;
            // Display the new form.
            newProcShiftCodeWin.Show(); 
        }
    }
}
