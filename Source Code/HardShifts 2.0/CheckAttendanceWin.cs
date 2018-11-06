using System;
using System.IO;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;


namespace HWAttendanceGrabber
{
    public partial class CheckAttendanceWin : Form
    {
        public CheckAttendanceWin()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            configFunction AttendanceConfig = new configFunction();
            string ls_atConfigPath = AttendanceConfig.getConfigFilePath();
            string ls_AttendancePath = AttendanceConfig.getAttendanceFilePath();

            
            int count = 0;
            int temp = 0;
            string filesList = null;

            ls_AttendancePath = ls_AttendancePath + "Attendance_2012" + "\\12_December 2012";

            DirectoryInfo attendanceDir = new DirectoryInfo(ls_AttendancePath);

            int fileCOunt = attendanceDir.GetFiles().Length;
            string[] fileName = new string[fileCOunt];

            foreach (FileInfo file in attendanceDir.GetFiles())
            {

                //MessageBox.Show(file.Name.ToString());                
                fileName[count] = file.Name.ToString();
                temp = count + 1;

                if(filesList == null)
                {                    
                    filesList = "\n\n" + "[" + temp + "] " + fileName[count];
                }
                else
                {
                    filesList = filesList + "\n" + "[" + temp + "] " + fileName[count];
                }                
                count++;                
            }

            if (filesList == null || filesList == "")
            {
                filesList = "[NONE] - Check your attendance directory path...";
            }

            MessageBox.Show("The following files are obtained: \n" +
                            "---------------------------------------\n\n" +
                            "FILE COUNT: " + fileCOunt + " FILES FOUND!" + "\n" +
                            "FILE NAMES: " + filesList + "\n\n" +
                            "[END]", "RESULT: ATTENDANCE FILE CHECK");                                                    
        }





        private void button1_Click(object sender, EventArgs e)
        {
            configFunction AttendanceConfig = new configFunction();
            string ls_atConfigPath = AttendanceConfig.getConfigFilePath();
            string ls_AttendancePath = AttendanceConfig.getAttendanceFilePath();            
            string ls_attTemplate = "C:\\consolidated_att_file.xlsx";

            string ls_firstFindAddress = null,
                   ls_currentFindAddress = null;


            int li_rowNumber = 0;
            int count = 0;
            int conAttRow = 4;
            int countAttendanceFile = 0;


            Excel._Application  newApplication = null;
            Excel._Workbook     newBook = null,
                                newBookTemplate = null;

            Excel._Worksheet    newSheet = null,
                                newSheetTemplate = null;

            Excel.Range          rngFirstFind = null,
                                rngCurrentFind = null;
                                //rngLastConAttendanceSheet = null; //the last range in the consolidated sheet template prior to opening a new ATT file    

            Excel.Range xl_badgeRange = null,
                                xl_nameRange = null,
                                xl_atDateRange = null,
                                xl_timeInRange = null,
                                xl_timeOutRange = null,
                                xl_totalHoursRange = null;
                                             
            DirectoryInfo attendanceDir = new DirectoryInfo(ls_AttendancePath);
            ShiftFunction ShiftCodeProcessor = new ShiftFunction();
            
   
            /*-- Get the FULL DIRECTORY of the attendance sheet for the given month --*/            
            foreach (DirectoryInfo folder in attendanceDir.GetDirectories())                        
            {                              
                if (folder.Name.Contains("2012"))                                                   
                {                    
                    ls_AttendancePath = ls_AttendancePath + folder.Name + "\\12_December 2012";    //Text in qoutes will be replaced with a variable name                
                }                
            }                        
            /*-----------------------------------------------------------------------**/

            string[] lsASFileArray = Directory.GetFiles(ls_AttendancePath, "*.xls");         

            newApplication = new Excel.Application();

            //Open the consolidated attendance template
            try
            {
                newBookTemplate = newApplication.Workbooks.Open(ls_attTemplate, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                newSheetTemplate = (Excel._Worksheet)newBookTemplate.Worksheets.get_Item(1);
                //newApplication.Visible = true;
            }
            catch
            {
                MessageBox.Show("ERROR: Consolidated attendance template cannot be found.");
            }

            ArrayList badgeNum = new ArrayList();
            configFunction configBadgeGet = new configFunction();


            
            //Loop through each attendance sheet                        
            foreach(string sheet in lsASFileArray)
            {                
                badgeNum = configBadgeGet.getBadgeArray(configBadgeGet.getBadgeFilePath(), "HWD");
          

                // Open the attendance sheet
                try
                {
                    newBook = newApplication.Workbooks.Open(sheet, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    newSheet = (Excel._Worksheet)newBook.Worksheets.get_Item(1);
                    countAttendanceFile++;
                    //newApplication.Visible = true;                    
                }
                catch
                {
                    MessageBox.Show("ERROR: Attendance sheet cannot be found.");
                    break;
                }


               //TemplateDateRangeEdit(newSheetTemplate, shiftPeriod, shiftYear);
                                                               
                foreach (string id in badgeNum)
                {                   
                    rngCurrentFind = newSheet.Cells.Find(id, System.Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, System.Type.Missing, System.Type.Missing);
                    ls_currentFindAddress = rngCurrentFind.get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
                                    
                    while (rngCurrentFind != null)
                    {
                        if (rngFirstFind == null)
                        {
                            rngFirstFind = rngCurrentFind;
                            ls_firstFindAddress = rngFirstFind.get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
                        }

                        else if (rngCurrentFind.get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing)
                        == rngFirstFind.get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing))
                        {
                            break;
                        }

                        li_rowNumber = rngCurrentFind.Row;

                        double badge = 0;
                        string name = null;
                        DateTime attendanceDate;
                        DateTime timeIN;
                        DateTime timeOUT;

                        double timeIN_TEMP;
                        double timeOUT_TEMP;
                        string TIME_IN;
                        string TIME_OUT;
                        string totalHours;

                        bool numCheck;
                        decimal num;

                        xl_badgeRange = newSheet.Cells[li_rowNumber, 2];
                        xl_nameRange = newSheet.Cells[li_rowNumber, 3];
                        xl_atDateRange = newSheet.Cells[li_rowNumber, 4];
                        xl_timeInRange = newSheet.Cells[li_rowNumber, 5];
                        xl_timeOutRange = newSheet.Cells[li_rowNumber, 6];
                        xl_totalHoursRange = newSheet.Cells[li_rowNumber, 7];

                        badge = xl_badgeRange.get_Value(System.Type.Missing);
                        name = xl_nameRange.get_Value(System.Type.Missing);
                        attendanceDate = xl_atDateRange.get_Value(System.Type.Missing);

                        if (xl_timeInRange.get_Value(System.Type.Missing) == null)
                        {
                            //timeIN = DateTime.Now;
                            TIME_IN = "";
                        }
                        else
                        {
                            timeIN_TEMP = xl_timeInRange.get_Value(System.Type.Missing);
                            //timeIN = DateTime.FromOADate(timeIN_TEMP);
                            TIME_IN = DateTime.FromOADate(timeIN_TEMP).ToShortTimeString();
                        }

                        if (xl_timeOutRange.get_Value(System.Type.Missing) == null)
                        {
                            //timeOUT = DateTime.Now;
                            TIME_OUT = "";
                        }
                        else
                        {
                            timeOUT_TEMP = xl_timeOutRange.get_Value(System.Type.Missing);
                            //timeOUT = DateTime.FromOADate(timeOUT_TEMP);
                            TIME_OUT = DateTime.FromOADate(timeOUT_TEMP).ToShortTimeString();
                        }

                        totalHours = xl_totalHoursRange.get_Value(System.Type.Missing).ToString();
                        numCheck = decimal.TryParse(totalHours.ToString(), out num);

                        /*
                        if (numCheck == false)
                        {                                                                                    
                            totalHours = "0";
                        }
                         * */

                        newSheetTemplate.Cells[conAttRow, 1] = badge.ToString();
                        newSheetTemplate.Cells[conAttRow, 2] = name;
                        newSheetTemplate.Cells[conAttRow, 3] = attendanceDate.ToString("yyyy/MM/dd");
                        newSheetTemplate.Cells[conAttRow, 4] = TIME_IN;
                        newSheetTemplate.Cells[conAttRow, 5] = TIME_OUT;
                        newSheetTemplate.Cells[conAttRow, 6] = totalHours;
                        
                        /*
                        else
                        {
                            //Write additional code here to continue populating consolidated attendance template 
                            //without opening another sheet
                        }
                        */

                        rngCurrentFind = newSheet.Cells.FindNext(rngCurrentFind);

                        xl_badgeRange = null;
                        xl_nameRange = null;
                        xl_atDateRange = null;
                        xl_timeInRange = null;
                        xl_timeOutRange = null;
                        xl_totalHoursRange = null;

                        conAttRow++;
                        //GET RANGE OF LAST ROW THAT WAS POPULATED...
                        //This will be the used to locate the
                    }
                    

                    /*--CLEAR the FIND variables to prepare for next BADGE number search--*/
                    rngCurrentFind = null;
                    rngFirstFind = null;
                    /*-----------------------------------------------------------------------*/

                  
                    if (count < badgeNum.Count - 2)
                    {
                        count++;
                    }
                    else
                    {
                        break;
                    }
                  
                }
                                
                Marshal.FinalReleaseComObject(newSheet);
                Marshal.FinalReleaseComObject(newBook);
                count = 0;
            }



            if (countAttendanceFile == 4) //time to save!
            {

                newSheetTemplate.Cells[2, 1] = "Consolidated Attendance For January 2013";

                //Generate a filename for the processed shift code document
                //System prompts to save the generated shift codes worksheet.
                //If user cancels, generated shift code worksheet will not be saved, and system terminates.             
                string lsSaveDate = String.Format("{0:MMM-dd-yyyy}", DateTime.Now);

                string saveFileName = "Consolidated Attendance for " + lsSaveDate +".xlsx";

                SaveFileDialog saveTemplateDialog = new SaveFileDialog();
                saveTemplateDialog.FileName = saveFileName;

                if (saveTemplateDialog.ShowDialog() == DialogResult.Cancel)
                {
                    MessageBox.Show("Save cancelled. Terminating program.");
                    return;
                }

                newBookTemplate.SaveAs(saveTemplateDialog.FileName, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

                //Close excel instances - save template, don't save shift code doc                                    
                newBookTemplate.Close(true, System.Type.Missing, System.Type.Missing);
                newApplication.Quit();


            }
            Marshal.FinalReleaseComObject(newApplication);
            newSheet = null;
            newSheetTemplate = null;
            newBook = null;
            newBookTemplate = null;
            newApplication = null;



            GC.WaitForPendingFinalizers();
            GC.Collect();
        }







        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void grabAttendanceWin_Load(object sender, EventArgs e)
        {
            //ddFrom.Enabled = false;
            //ddTo.Enabled = false;

            /*
            string defaultSection;
            defaultSection = "[HW]";
            dd_Section.SelectedItem = defaultSection;


            //1.) POPULATE YEAR DDLIST STARTING FROM 2011 TO CURRENT YEAR            
            int listCurrentYear = DateTime.Now.Year;
            int listStartYear = 2011;

            do
            {
                dd_Year.Items.Add(listStartYear);
                listStartYear++;
            }
            while (listStartYear <= listCurrentYear);
            //--END OF (1)

            //2.) SET YEAR TO CURRENT YEAR
            dd_Year.SelectedItem = listCurrentYear;
            //--END of (2)
             * */

        }

        private void labelCheckFile_Click(object sender, EventArgs e)
        {

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            
        }


    }
}
