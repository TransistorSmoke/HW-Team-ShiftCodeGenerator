using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;


namespace HWAttendanceGrabber
{
    public partial class ProcessShiftCodeWin : Form
    {
        public ProcessShiftCodeWin()
        {
            InitializeComponent();            
        }


        private void ProcessShiftCodeWin_Load(object sender, EventArgs e)
        {
          

            //1.) POPULATE YEAR DDLIST STARTING FROM 2011 TO CURRENT YEAR
            /*
            int listCurrentYear = DateTime.Now.Year;
            int listStartYear = 2011;

            do
            {
                ddShiftYear.Items.Add(listStartYear);
                listStartYear++;
            }
            while (listStartYear <= listCurrentYear+2);
            //--END OF (1)

            //2.) SET YEAR TO CURRENT YEAR
            ddShiftYear.SelectedItem = listCurrentYear;
            //--END of (2)
             */
        }

        public int ValidateArguments(string file, string period, string section)
        {
            string  filename = "",
                    periodStat = "",
                    sectionStat = "",                    
                    eightHourDiffStat = "";
          
            string valMessage = "ERRORS FOUND:\n-----------------\n\n";

            if (file == "" || file == null)
            {
                filename = null;  
                valMessage = valMessage + "- Please select DAILY ATTENDANCE FILE.\n";                
            }

            if(period == "" || period == null)
            {
                periodStat = null;  
                valMessage = valMessage + "- Please select ATTENDANCE PERIOD.\n";                
            }

            if (section == "" || section == null)
            {
                sectionStat = null;
                valMessage = valMessage + "- Please choose your SECTION.\n";                
            }

            if (radioButtonYes.Checked == true)
            {
                if (ddFrom.Text == null || ddTo.Text == null || ddFrom.Text == "" || ddTo.Text == "")
                {
                    eightHourDiffStat = null;
                    valMessage = valMessage + "- 8-Hour Shift period checked. Please supply appropriate DATE RANGE.\n";                    
                }

                else if (ddFrom.Text != "" || ddTo.Text != "")
                {
                    int eightHrDayDiff = Convert.ToInt32((Convert.ToDateTime(ddFrom.Text) - Convert.ToDateTime(ddTo.Text)).TotalDays);
                    if (eightHrDayDiff >= 0)
                    {
                        valMessage = valMessage + "- 8-Hour Shift: DATE FROM is greater than DATE TO. Please correct.\n";
                        eightHourDiffStat = null;
                    }
                }
            }
            
            /*
            if(radioButtonYes.Enabled == true)
            {
                if ((dateFrom_8hours == null || dateTo_8hours == null) || (dateFrom_8hours == "" || dateTo_8hours == ""))
                {
                    check8Hours = 0;
                    valMessage = valMessage + "- 8-Hour Shift period checked. Please supply appropriate DATE RANGE.\n";
                }

                else if((dateFrom_8hours != "" || dateTo_8hours != "") && (dateFrom_8hours != null || dateTo_8hours != null))
                {
                    int eightHrDayDiff = Convert.ToInt32((Convert.ToDateTime(dateFrom_8hours) - Convert.ToDateTime(dateTo_8hours)).TotalDays);
                    if (eightHrDayDiff >= 0)
                    {
                        valMessage = valMessage + "- 8-Hour Shift: DATE FROM is greater than DATE TO. Please correct.\n";
                        eightHourDiffStat = null;
                    }
                }
            }
             */


            if (filename == null || periodStat == null || sectionStat == null || eightHourDiffStat == null)
            {
                MessageBox.Show(valMessage);
                return 0;
            }
            else
            {
                return 1;
            }
        }

        private void btnGetData_Click(object sender, EventArgs e)
        {
            int empIDLen = 0;
            int li_rowNumber = 0;
            int li_badgeSCRow = 0;
            int tempRow = 0;
            int li_column = 0;
            int col = 0;
            int templatePopulate_ROW = 0;
            int templatePopulate_COLUMN = 0;
            int validationStat;
            int hourCheck = 1;
            
            string consolidatedAttendanceSheet = fileNameBox.Text;
            string ls_AttendanceFileName = null;
            string ls_attendanceYear = null;
            //string shiftCodeTemplate = "C:\\SW_template.xlsx";
            //string shiftCodeTemplate = "C:\\HW_template.xlsx";
            string shiftCodeTemplate = null;
            string ls_shiftPeriod = null;
            string SECTION = null, ls_SECTION = null;
            string TIME_IN = "",
                     TIME_OUT = "",
                     SHIFT_DATE,
                     template_RefShiftDate = null,
                     dateFrom_8hours = null,
                     dateTo_8hours = null;


            DateTime temp_SHIFT_DATE;

            double temp_TIME_IN,
                   temp_TIME_OUT;


            string  ls_firstFindAddress = null,
                    ls_currentFindAddress = null,
                    ls_firstFindSCTemplateAddress = null,
                    ls_currentFindSCTemplateAddress = null,
                    ls_badgeFindSC = null,
                    ls_prevDateRangeAddress = null,
                    ls_currentDateRangeAddress = null,
                    ls_BadgeAttendanceDate = null,
                    ls_thisDaysShiftCode = null;


            Excel._Application  newExcelApp = null;
            Excel._Workbook     newConsolidatedBook = null,
                                newSCBook = null;

            Excel._Worksheet    newConsolidatedSheet = null,
                                newSCSheet = null;

            Excel.Range rngFirstFind = null,
                                rngCurrentFind = null,
                                rngBadgeSCFind = null,
                                rngFirstFindSCTemplate = null,
                                rngCurrentFindSCTemplate = null,
                                rngTemplateDateSearch = null,
                                rngDefaultShiftCode = null;

            Excel.Range         rngBadgeAttendanceDate = null;
                                    
            ls_AttendanceFileName = fileNameBox.Text;
            ls_shiftPeriod = ddShiftPeriod.Text;
            ls_SECTION = ddSection.Text;
            
            if(radioButtonNo.Checked)
            {
                
                dateFrom_8hours = null;
                dateTo_8hours = null;
                hourCheck = 0;
            }
            else if(radioButtonYes.Checked)

            {                
                dateFrom_8hours = ddFrom.Text;
                dateTo_8hours = ddTo.Text;
                hourCheck = 1;
            }

            /*Validate inputs.... */
            validationStat = ValidateArguments(ls_AttendanceFileName, ls_shiftPeriod, ls_SECTION);

            if (validationStat == 0)
            {
                return;
            }

            ShiftFunction SHIFTCODE_GEN_INIT = new ShiftFunction();
            
            shiftCodeTemplate = SHIFTCODE_GEN_INIT.getTemplatePath(ls_SECTION);
            

                                                                                                                                                            
            newExcelApp = new Excel.Application();

            try
            {
                newSCBook = newExcelApp.Workbooks.Open(shiftCodeTemplate, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                newSCSheet = (Excel._Worksheet)newSCBook.Worksheets.get_Item(1);
                //newExcelApp.Visible = true;
            }

            catch
            {
                MessageBox.Show("ERROR: SHIFT CODE TEMPLATE cannot be opened.");
                return;
            }

            try
            {
                newConsolidatedBook = newExcelApp.Workbooks.Open(consolidatedAttendanceSheet, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                newConsolidatedSheet = (Excel._Worksheet)newConsolidatedBook.Worksheets.get_Item(1);
                //newExcelApp.Visible = true;
            }

            catch
            {
                MessageBox.Show("ERROR: ATTENDANCE SHEET cannot be opened.");
                return;
            }
            
            ls_attendanceYear = yearTextBox.Text;
            //ls_shiftPeriod = ddShiftPeriod.Text;

            newSCSheet.Cells[3, 1] = "Attendance Period: " + ddShiftPeriod.Text;

            TemplateDateRangeEdit(newSCSheet, ls_shiftPeriod, ls_attendanceYear);

                                  
            ArrayList activeEmpID = new ArrayList();
            configFunction configFn = new configFunction();

            if (ls_SECTION == "HARDWARE DEVELOPMENT")
            {
                SECTION = "HWD";
            }

            else if (ls_SECTION == "SOFTWARE DEVELOPMENT")
            {
                SECTION = "SWD";
            }

            else if (ls_SECTION == "SOFTWARE VALIDATION")
            {
                SECTION = "SWV";
            }

            else if (ls_SECTION == "ELECTROMECHANICAL")
            {
                SECTION = "EM";
            }

            else if (ls_SECTION == "LAB, PM AND PQ")
            {
                SECTION = "LAB_PM_PQ";
            }

            else
            {
                SECTION = "ALL";
            }
            activeEmpID = getIDNum(configFn.getBadgeFilePath(), SECTION);
            empIDLen = activeEmpID.Count - 1;

            for (int column = 0; column <= 14; column++)
            {            
                foreach (string id in activeEmpID)
                {                   
                   if (id == null || id == "") //if id does not contain value anymore, search is done. Exit and close.
                   {                       
                       break;
                   }

                   string ls_TEMPORARY_badgeInTemplate_ADDRESS = null;
                   
                   rngTemplateDateSearch = newSCSheet.Cells[5, 5+column];
                   template_RefShiftDate = string.Format("{0:d-MMM-yy}", rngTemplateDateSearch.get_Value(System.Type.Missing));
                                   
                   rngCurrentFindSCTemplate = newSCSheet.Cells.Find(id, System.Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, System.Type.Missing, System.Type.Missing);

                   if (rngCurrentFindSCTemplate == null)
                   {
                       
                       continue;
                   }

                   else
                   {
                        ls_TEMPORARY_badgeInTemplate_ADDRESS = rngCurrentFindSCTemplate.get_Address(System.Type.Missing);
                   }

                   templatePopulate_ROW = rngCurrentFindSCTemplate.Row;
                   templatePopulate_COLUMN = 5+column;

                   //Find the badge number in the Consolidated Attendance Sheet
                   rngCurrentFind = newConsolidatedSheet.Cells.Find(id, System.Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, System.Type.Missing, System.Type.Missing);

                    
                   if (rngCurrentFind == null) //which means the employee with searched badge number HAS RESIGNED (still present in the badge list in the INI file but not in the attendance anymore)
                   {
                       continue;
                   }


                   else
                   {

                       ls_currentFindAddress = rngCurrentFind.get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);

                       //Badge number found!
                       while (rngCurrentFind != null)
                       {
                           //First find here! Save the range of the first found cell
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

                           string ls_tempAttendanceDate = null;
                           rngBadgeAttendanceDate = newConsolidatedSheet.Cells[li_rowNumber, 3];
                           ls_tempAttendanceDate = string.Format("{0:d-MMM-yy}", rngBadgeAttendanceDate.get_Value(System.Type.Missing));

                           //SHIFT_DATE = ls_tempAttendanceDate.Substring(0, 8);
                           SHIFT_DATE = ls_tempAttendanceDate;

                          
                           //Upon finding badge in the attendance sheet, check attendance date...
                           if (template_RefShiftDate != SHIFT_DATE)
                           {
                               int daysLapse = Convert.ToInt32((Convert.ToDateTime(template_RefShiftDate) - Convert.ToDateTime(SHIFT_DATE)).TotalDays);
                              
                               if (daysLapse <= 0)
                               {
                                   ls_thisDaysShiftCode = ConvertToShiftCode(Convert.ToDateTime(template_RefShiftDate), TIME_IN, "9.6hours");

                                   if (Convert.ToDateTime(template_RefShiftDate).DayOfWeek.ToString() == "Saturday")
                                   {
                                       newSCSheet.Cells[templatePopulate_ROW, templatePopulate_COLUMN] = "SAT";
                                   }

                                   else if (Convert.ToDateTime(template_RefShiftDate).DayOfWeek.ToString() == "Sunday")
                                   {
                                       newSCSheet.Cells[templatePopulate_ROW, templatePopulate_COLUMN] = "R";
                                   }

                                   else
                                   {
                                       newSCSheet.Cells[templatePopulate_ROW, templatePopulate_COLUMN] = "E10";
                                   }

                                   rngDefaultShiftCode = newSCSheet.Cells[templatePopulate_ROW, templatePopulate_COLUMN];
                                   rngDefaultShiftCode.Font.Bold = true;
                                   rngDefaultShiftCode = null;
                                   break;
                               }

                               rngCurrentFind = newConsolidatedSheet.Cells.FindNext(rngCurrentFind);
                           }

                           else
                           {
                               Excel.Range rngSHIFT_DATE = newConsolidatedSheet.Cells[li_rowNumber, 3],
                                           rngTYPE_CHECK = newConsolidatedSheet.Cells[li_rowNumber, 4],
                                           rngTIME_IN = null;

                               string TYPE_CHECK = rngTYPE_CHECK.get_Value(System.Type.Missing);

                               try
                               {
                                   temp_SHIFT_DATE = rngSHIFT_DATE.get_Value(System.Type.Missing);
                                   SHIFT_DATE = string.Format("{0:d-MMM-yy}", temp_SHIFT_DATE);
                               }
                               catch
                               {
                                   SHIFT_DATE = "";
                               }

                               try
                               {
                                   if (TYPE_CHECK.ToUpper() == "IN")
                                   {
                                       rngTIME_IN = newConsolidatedSheet.Cells[li_rowNumber, 5];
                                       temp_TIME_IN = rngTIME_IN.get_Value(System.Type.Missing);
                                       TIME_IN = DateTime.FromOADate(temp_TIME_IN).ToShortTimeString();
                                   }
                               }
                               catch
                               {
                                   TIME_IN = "";
                               }

                               ls_thisDaysShiftCode = ConvertToShiftCode(Convert.ToDateTime(SHIFT_DATE), TIME_IN, "9.6hours");
                               newSCSheet.Cells[templatePopulate_ROW, templatePopulate_COLUMN] = ls_thisDaysShiftCode;

                               if (ls_thisDaysShiftCode == "L10" || ls_thisDaysShiftCode == "L12" || ls_thisDaysShiftCode == "SAT" || ls_thisDaysShiftCode == "R")
                               {
                                   rngDefaultShiftCode = newSCSheet.Cells[templatePopulate_ROW, templatePopulate_COLUMN];
                                   rngDefaultShiftCode.Font.Bold = true;
                                   rngDefaultShiftCode = null;
                                   break;
                               }
                               else
                               {
                                   break;
                               }
                               
                           }

                           //rngCurrentFind = newConsolidatedSheet.Cells.FindNext(rngCurrentFind);
                           ls_currentFindAddress = rngCurrentFind.get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);

                       }

                   }
                  
                   rngCurrentFind = null;
                   rngFirstFind = null;

                }
            }

            rngCurrentFindSCTemplate = null;
            rngFirstFindSCTemplate = null;

            //Generate a filename for the processed shift code document
            //System prompts to save the generated shift codes worksheet.
            //If user cancels, generated shift code worksheet will not be saved, and system return 0 and terminates.             
            string lsSaveDate = String.Format("{0:MMMddyyyy}", DateTime.Now);

            string saveFileName = ls_SECTION + "_Shifts_" + ls_shiftPeriod + "_" + lsSaveDate + ".xls";

            SaveFileDialog saveTemplateDialog = new SaveFileDialog();
            saveTemplateDialog.FileName = saveFileName;

            if (saveTemplateDialog.ShowDialog() == DialogResult.Cancel)
            {
                MessageBox.Show("Save cancelled. Terminating program.");
            }

            else
            {
                MessageBox.Show("SHIFT CODE GENERATION SUCCESSFUL. \n\nFILE NAME: " + saveFileName);
            }

            

            newSCBook.SaveAs(saveTemplateDialog.FileName, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
            newExcelApp.Quit();

            //Close Excel, release objects, do garbage collection
            
            Marshal.FinalReleaseComObject(newConsolidatedSheet);
            Marshal.FinalReleaseComObject(newConsolidatedBook);
            Marshal.FinalReleaseComObject(newSCSheet);
            Marshal.FinalReleaseComObject(newSCBook);
            Marshal.FinalReleaseComObject(newExcelApp);
            
            newConsolidatedBook = null;
            newSCBook = null;
            newConsolidatedSheet = null;
            newSCSheet = null;
            newExcelApp = null;

            GC.WaitForPendingFinalizers();
            GC.Collect();              
            
        }



        public ArrayList getIDNum(string iniSection, string cesection)
        {

            string id;
            int idCount = 0;
            ArrayList empActive = new ArrayList();

            configFunction configFn = new configFunction();
            configFile config = new configFile(configFn.getBadgeFilePath());

            //RETRIEVE REGISTERED BADGE NUMBERS FROM CONFIG.INI
            do
            {
                id = config.Read(cesection, "idNum[" + idCount + "]");
                empActive.Add(id);
                idCount++;
            }
            while (id != "");
            //END OF ID RETRIEVAL

            return empActive;       //RETURNS AN ARRAYLIST OF BADGE NUMBERS
        }


        public string MONTH_STRING_TO_INT(string month)
        {
            if (month == "Jan")
            {
                return "01";
            }
            else if (month == "Feb")
            {
                return "02";
            }
            else if (month == "Mar")
            {
                return "03";
            }
            else if (month == "Apr")
            {
                return "04";
            }
            else if (month == "May")
            {
                return "05";
            }
            else if (month == "Jun")
            {
                return "06";
            }
            else if (month == "Jul")
            {
                return "07";
            }
            else if (month == "Aug")
            {
                return "08";
            }
            else if (month == "Sep")
            {
                return "09";
            }
            else if (month == "Oct")
            {
                return "10";
            }
            else if (month == "Nov")
            {
                return "11";
            }
            else if (month == "Dec")
            {
                return "12";
            }
            else
            {
                return "";
            }

        }




        //Converts/Populates the cells E5 to S5/T5 in shift code template sheet with appropriate dates based on shift period
        public void TemplateDateRangeEdit(Excel._Worksheet workSheet, string shiftPeriod, string shiftYear)
        {
            string lsEndPeriod, lsYear;
            int liDay, cellCount = 15;

            liDay = DateTime.Now.Day; //Gets day            
            lsYear = shiftYear;

            lsEndPeriod = shiftPeriod.Substring(7);

            if (Convert.ToInt32(lsEndPeriod) >= 28)     //Month ends in 28, 29, 30 or 31
            {
                for (cellCount = 16; cellCount <= (Convert.ToInt32(lsEndPeriod)); cellCount++)
                {
                    workSheet.Cells[5, cellCount - 11] = cellCount + "-" + shiftPeriod.Substring(0, 3) + "-" + lsYear;
                    //MessageBox.Show(cellCount + "-" + shiftPeriod.Substring(0, 3) + "-" + lsYear);                        
                }
            }

            else    //shift period is from day 01-15
            {
                for (cellCount = 1; cellCount <= 15; cellCount++)
                {
                    if (cellCount < 10)
                    {
                        workSheet.Cells[5, cellCount + 4] = "0" + cellCount + "-" + shiftPeriod.Substring(0, 3) + "-" + lsYear;
                    }

                    else
                    {
                        workSheet.Cells[5, cellCount + 4] = cellCount + "-" + shiftPeriod.Substring(0, 3) + "-" + lsYear;
                    }
                }
            }

        }

        private void SelectFileBtn_Click(object sender, EventArgs e)
        {
            string filename = null;
            string valYear = null; 


            OpenFileDialog result = new OpenFileDialog();

            if(result.ShowDialog() ==  DialogResult.OK)
            {
                fileNameBox.Text = result.FileName;
                filename = Path.GetFileName(fileNameBox.Text);
                
                if (string.IsNullOrEmpty(filename))
                {
                    MessageBox.Show("No attendance sheet selected.");
                }

                else
                {
                    string yeartest = "201";
                    string ls_attendanceYear = null;
                    string ls_extractedMonth = null;
                    int li_found;



                    /* Extract file name from the attendance sheet path and retrieve MONTH info */
                    ls_extractedMonth = ExtractMonthFromFileName(filename);

                    if(ls_extractedMonth == null) //PROMPT TO USE THE PROPER ATTENDANCE FILE NAMING AND TEMPLATE
                    {
                        MessageBox.Show("Please check attendance file. \n\nIMPORTANT\n\n 1.) DO NOT CHANGE THE ATTENDANCE FILE NAMING NOMENCLATURE.\n 2.) Format used is 'Timelog yyyy-dd <month>.xlsx - eg. Timelog 2013-'.");
                        return;
                    }
                    
                    /* Retrieve YEAR info from filename and populate YEAR TextBox with year extracted from the filename */
                    li_found = filename.IndexOf(yeartest);
                    ls_attendanceYear = filename.Substring(li_found, 4);                    
                    yearTextBox.Text = ls_attendanceYear;

                    PopulateShiftPeriod(ls_extractedMonth, ls_attendanceYear);
                                    
                }
            
            }



        }


        //This function generates the shift code and posts it to the template      
        public string ConvertToShiftCode(DateTime refDatetime, string timeIN, string WorkHours)
        {
            string timeStamp;
            DateTime dtTimeStamp;

            if (refDatetime.DayOfWeek.ToString() == "Saturday")
            {
                return "SAT";
            }

            else if (refDatetime.DayOfWeek.ToString() == "Sunday")
            {
                //MessageBox.Show("I found a SUNDAY!");
                return "R";
            }

            else if ((refDatetime.DayOfWeek.ToString() == "" || refDatetime.DayOfWeek.ToString() == null) && (refDatetime.DayOfWeek.ToString() != "Saturday") || refDatetime.DayOfWeek.ToString() == "Sunday")
            {
                return "E10 - Blank";
            }

            else
            {
                if (timeIN.ToString() == ("00:00:00") || timeIN.ToString() == "") //If TIME IN record is NULL or empty string, halt processing and return empty string
                {

                    return "E10";       /* ADDED ON FEBRUARY 24, 2012
                                         * Suggested by Charity. In case the engineer has no login info due to 
                                         * failure to swipe in, Logbox entry, or leave of absence, the system should generate
                                         * E10 as a default shift code.                                        
                                         */
                }

                else
                {
                    timeStamp = string.Format("{0:HH:mm:ss}", timeIN);
                    dtTimeStamp = Convert.ToDateTime(timeStamp);

                    if (WorkHours == "9.6hours")
                    {
                        //if ((dtTimeStamp >= Convert.ToDateTime("06:00 AM") && dtTimeStamp <= Convert.ToDateTime("09:00 AM")))
                        //if ((dtTimeStamp >= Convert.ToDateTime("06:00 AM") && dtTimeStamp <= Convert.ToDateTime("09:00 AM")) || (dtTimeStamp > Convert.ToDateTime("09:00 AM") && dtTimeStamp <= Convert.ToDateTime("09:38 AM")))
                        if ((dtTimeStamp >= Convert.ToDateTime("06:00 AM") && dtTimeStamp <= Convert.ToDateTime("09:00 AM")))
                        {
                            return "E10";
                        }

                        else if (dtTimeStamp > Convert.ToDateTime("09:00 AM") && dtTimeStamp <= Convert.ToDateTime("09:38 AM"))
                        {

                            return "L10";   //LATE - Closest shift is E10
                        }

                        else if (dtTimeStamp >= Convert.ToDateTime("05:45 AM") && dtTimeStamp <= Convert.ToDateTime("07:15 AM"))
                        {
                            return "E11";
                        }

                        //else if ((dtTimeStamp >= Convert.ToDateTime("10:15 AM") && dtTimeStamp <= Convert.ToDateTime("01:15 PM")) || (dtTimeStamp > Convert.ToDateTime("09:30 AM") && dtTimeStamp < Convert.ToDateTime("10:15 AM")))
                        //else if ((dtTimeStamp >= Convert.ToDateTime("10:15 AM") && dtTimeStamp <= Convert.ToDateTime("01:15 PM")) || (dtTimeStamp > Convert.ToDateTime("09:38 AM") && dtTimeStamp < Convert.ToDateTime("10:15 AM")))
                        else if ((dtTimeStamp >= Convert.ToDateTime("10:15 AM") && dtTimeStamp <= Convert.ToDateTime("01:15 PM")))
                        {
                            return "E12";
                        }

                        else if (dtTimeStamp > Convert.ToDateTime("09:38 AM") && dtTimeStamp < Convert.ToDateTime("10:15 AM"))
                        {

                            return "L12"; //LATE - Closest shift is E12
                        }


                        else if (dtTimeStamp >= Convert.ToDateTime("09:30 PM") && dtTimeStamp <= Convert.ToDateTime("10:00 PM"))
                        {
                            return "E13";
                        }

                        else if ((dtTimeStamp >= Convert.ToDateTime("07:00 PM") && dtTimeStamp <= Convert.ToDateTime("10:00 PM")) || (dtTimeStamp > Convert.ToDateTime("03:30 PM") && dtTimeStamp < Convert.ToDateTime("07:00 PM")))
                        {
                            return "E14";
                        }

                        /*
                         * SHIFT CODE E17 REMOVED, effective April 16, 2012
                        else if (dtTimeStamp >= Convert.ToDateTime("09:00 AM") && dtTimeStamp <= Convert.ToDateTime("09:30 AM"))
                        {
                            return "E17";
                        }
                         * */

                        /*
                        else if (dtTimeStamp >= Convert.ToDateTime("07:30 PM") && dtTimeStamp <= Convert.ToDateTime("08:00 PM"))
                        {
                            return "E65";
                        }

                        else if ((dtTimeStamp >= Convert.ToDateTime("02:30 PM") && dtTimeStamp <= Convert.ToDateTime("03:30 PM")) || (dtTimeStamp > Convert.ToDateTime("01:15 PM") && dtTimeStamp < Convert.ToDateTime("02:30 PM")))
                        {
                            return "E66";
                        }

                        else if (dtTimeStamp >= Convert.ToDateTime("06:30 AM") && dtTimeStamp <= Convert.ToDateTime("08:30 AM"))
                        {
                            return "E80";
                        }
                        */

                        else if (dtTimeStamp >= Convert.ToDateTime("04:45 AM") && dtTimeStamp <= Convert.ToDateTime("06:45 AM") || (dtTimeStamp > Convert.ToDateTime("02:00 AM") && dtTimeStamp <= Convert.ToDateTime("04:45 AM")))
                        {
                            return "E81";
                        }

        
                        /*        
                        else if (dtTimeStamp >= Convert.ToDateTime("10:45 AM") && dtTimeStamp <= Convert.ToDateTime("12:45 PM"))
                        {
                            return "E82";
                        }
                        */

                        else
                        {
                            return "";
                        }
                    }

                    else        //WorkHours == "8hours". Do 8-hour shift code processing here
                    {
                        if (dtTimeStamp >= Convert.ToDateTime("11:45 AM") && dtTimeStamp <= Convert.ToDateTime("12:15 PM"))
                        {
                            return "E15";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("06:00 AM") && dtTimeStamp <= Convert.ToDateTime("06:30 AM"))
                        {
                            return "E21";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("02:00 PM") && dtTimeStamp <= Convert.ToDateTime("02:30 PM"))
                        {
                            return "E22";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("10:00 PM") && dtTimeStamp <= Convert.ToDateTime("10:30 PM"))
                        {
                            return "E23";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("07:00 AM") && dtTimeStamp <= Convert.ToDateTime("08:00 AM"))
                        {
                            return "E40";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("05:15 AM") && dtTimeStamp <= Convert.ToDateTime("06:15 AM"))
                        {
                            return "E41";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("05:45 AM") && dtTimeStamp <= Convert.ToDateTime("06:15 AM"))
                        {
                            return "E61";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("01:45 PM") && dtTimeStamp <= Convert.ToDateTime("02:15 PM"))
                        {
                            return "E62";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("09:45 PM") && dtTimeStamp <= Convert.ToDateTime("10:15 PM"))
                        {
                            return "E63";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("07:00 PM") && dtTimeStamp <= Convert.ToDateTime("10:00 PM"))
                        {
                            return "E64";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("05:30 AM") && dtTimeStamp <= Convert.ToDateTime("08:30 AM"))
                        {
                            return "E95";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("07:30 AM") && dtTimeStamp <= Convert.ToDateTime("09:00 AM"))
                        {
                            return "E50";
                        }

                        if (dtTimeStamp >= Convert.ToDateTime("09:00 AM") && dtTimeStamp <= Convert.ToDateTime("12:00 PM"))
                        {
                            return "E51";
                        }

                        else
                        {
                            return "";
                        }

                    }


                }
            }
        }

        //Converts column number to representative integer
        public int colToInt(string s)
        {
            int colInt;
            char result = Convert.ToChar(s.Substring(1, 1));            
            colInt = Convert.ToInt32(result) - 64;
            return colInt;
        }

        //Retrieves MONTH data from the attendance filename
        public string ExtractMonthFromFileName(string filename)
        {


            if (filename.IndexOf("Jan", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Jan";
            }
            else if(filename.IndexOf("Feb", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Feb";
            }
                
            else if(filename.IndexOf("Mar", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Mar";
            }

            else if (filename.IndexOf("Apr", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Apr";
            }

            else if (filename.IndexOf("May", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "May";
            }

            else if (filename.IndexOf("Jun", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Jun";
            }

            else if (filename.IndexOf("Jul", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Jul";
            }

            else if (filename.IndexOf("Aug", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Aug";
            }

            else if (filename.IndexOf("Sep", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Sep";
            }

            else if (filename.IndexOf("Oct", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Oct";
            }

            else if (filename.IndexOf("Nov", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Nov";
            }

            else if (filename.IndexOf("Dec", StringComparison.OrdinalIgnoreCase) != -1)
            {
                return "Dec";
            }

            else
            {
                return null;
            }

        }

        void PopulateShiftPeriod(string month, string year)
        {
            string ls_1stQuincena = null;
            string ls_2ndQuincena = null;

            ddShiftPeriod.Items.Clear();

            ls_1stQuincena = month + " 01-15";

            if (month == "Jan" || month == "Mar" || month == "May" || month == "Jul" || month == "Aug" || month == "Oct" || month == "Dec")
            {
                ls_2ndQuincena = month + " 16-31";
            }

            else if (month == "Feb")
            {
                if ((Convert.ToInt32(year) % 4) == 0) //It's a leap year!
                {
                    ls_2ndQuincena = month + " 16-29";
                }
                else
                {
                    ls_2ndQuincena = month + " 16-28";
                }
            }

            else
            {
                ls_2ndQuincena = month + " 16-30";
            }

            ddShiftPeriod.Items.Add(ls_1stQuincena);
            ddShiftPeriod.Items.Add(ls_2ndQuincena);
        }

        void Populate_8Hr_Period()
        {
            string month;
            int period_start, period_end;

            ddFrom.Items.Clear();
            ddTo.Items.Clear();

            month = ddShiftPeriod.Text.Substring(0,3); 
            period_start = Convert.ToInt32(ddShiftPeriod.Text.Substring(4, 2));
            period_end = Convert.ToInt32(ddShiftPeriod.Text.Substring(7, 2));

            for (int i = period_start; i <= period_end; i++)
            {
                    ddFrom.Items.Add(month + " " + i.ToString());
            }

            for (int i = period_start; i <= period_end; i++)
            {
                    ddTo.Items.Add(month + " " + i.ToString());
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void radioButtonNo_CheckedChanged(object sender, EventArgs e)
        {
            ddFrom.Enabled = false;
            ddTo.Enabled = false;
            
        }

        private void radioButtonYes_CheckedChanged(object sender, EventArgs e)
        {
            string month;
            string period;
            int period_start, period_end;

            ddFrom.Enabled = true;
            ddTo.Enabled = true;
            
            if (ddShiftPeriod.Text == null || ddShiftPeriod.Text == "")
            {
                MessageBox.Show("Please select a SHIFT PERIOD first.");
            }

            else
            {
                Populate_8Hr_Period();
                /*
                ddFrom.Items.Clear();
                ddTo.Items.Clear();

                month = ddShiftPeriod.Text.Substring(0,3); 
                period_start = Convert.ToInt32(ddShiftPeriod.Text.Substring(4, 2));
                period_end = Convert.ToInt32(ddShiftPeriod.Text.Substring(7, 2));

                for (int i = period_start; i <= period_end; i++)
                {
                    ddFrom.Items.Add(month + " " + i.ToString());
                }

                for (int i = period_start; i <= period_end; i++)
                {
                    ddTo.Items.Add(month + " " + i.ToString());
                }
                */
            }            
        }

        private void ddShiftPeriod_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (radioButtonYes.Enabled == true)
            {
                Populate_8Hr_Period();
            }
        }
   
    }
}
