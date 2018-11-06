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
using System.Runtime.InteropServices;
using System.IO;

namespace HWAttendanceGrabber
{
    class ShiftFunction
    {
        public void GenerateCSV(string attendanceFilePath, string year) //, string period)
        {
            //string ls_atFileNameTest;
            //ls_atFileNameTest = "C:\\Daily Attendance\\Attendance_2012\\12_December 2012\\Dec01-9_12.xlsx";
            
            string ls_atFileNameTest;

            //ls_atFileNameTest = "C:\\Daily Attendance\\Attendance_2012\\12_December 2012\\Dec01-9_12.xlsx";
            ls_atFileNameTest = attendanceFilePath;
            string ls_attTemplate = "C:\\consolidated_att_file.xlsx";      

            string ls_firstFindAddress = null,
                   ls_currentFindAddress = null;
                   

            int li_rowNumber = 0;
            int count = 0;
            int conAttRow = 4;
            
            
            Excel._Application  newApplication = null;
            Excel._Workbook newBook = null,
                                newBookTemplate = null;
            
            Excel._Worksheet    newSheet = null,
                                newSheetTemplate = null;
            
            Excel.Range         rngFirstFind = null,
                                rngCurrentFind = null;

            Excel.Range         xl_badgeRange = null,
                                xl_nameRange = null,
                                xl_atDateRange = null,
                                xl_timeInRange = null,
                                xl_timeOutRange = null,
                                xl_totalHoursRange = null;
            

            newApplication = new Excel.Application();            

            ArrayList badgeNum = new ArrayList();
            configFunction configBadgeGet = new configFunction();            

            badgeNum = configBadgeGet.getBadgeArray(configBadgeGet.getBadgeFilePath(), "HWD");

            

                      
            /*--------------OPEN ATTENDANCE SHEET-------------*/
            try
            {
                newBook = newApplication.Workbooks.Open(ls_atFileNameTest, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                newSheet = (Excel._Worksheet)newBook.Worksheets.get_Item(1);
                //newApplication.Visible = true;
            }
            catch
            {
                MessageBox.Show("ERROR: Attendance sheet cannot be found.");                
            }
            /*------------------------------------------------*/

            /*----------OPEN CONSOLIDATED ATTENDANCE TEMPLATE -------*/

            try
            {
                newBookTemplate = newApplication.Workbooks.Open(ls_attTemplate, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                newSheetTemplate = (Excel._Worksheet)newBookTemplate.Worksheets.get_Item(1);
                newApplication.Visible = true;
            }
            catch
            {
                MessageBox.Show("ERROR: Consolidated attendance template cannot be found.");
            }

            /*------------------------------------------------*/

            



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
                        timeIN = DateTime.Now;
                    }
                    else
                    {
                        timeIN_TEMP = xl_timeInRange.get_Value(System.Type.Missing);
                        timeIN = DateTime.FromOADate(timeIN_TEMP);
                    }

                    if (xl_timeOutRange.get_Value(System.Type.Missing) == null)
                    {
                        timeOUT = DateTime.Now;
                    }
                    else
                    {
                        timeOUT_TEMP = xl_timeOutRange.get_Value(System.Type.Missing);
                        timeOUT = DateTime.FromOADate(timeOUT_TEMP);
                    }

                    totalHours = xl_totalHoursRange.get_Value(System.Type.Missing).ToString();
                    numCheck = decimal.TryParse(totalHours.ToString(), out num);

                    if (numCheck == false)
                    {
                        totalHours = "0";
                    }

                    newSheetTemplate.Cells[conAttRow, 1] = badge.ToString();
                    newSheetTemplate.Cells[conAttRow, 2] = name;
                    newSheetTemplate.Cells[conAttRow, 3] = attendanceDate.ToString("yyyy/MM/dd");
                    newSheetTemplate.Cells[conAttRow, 4] = timeIN.ToShortTimeString();
                    newSheetTemplate.Cells[conAttRow, 5] = timeOUT.ToShortTimeString();
                    newSheetTemplate.Cells[conAttRow, 6] = totalHours;
                

                    rngCurrentFind = newSheet.Cells.FindNext(rngCurrentFind); 

                    xl_badgeRange = null;
                    xl_nameRange = null;
                    xl_atDateRange = null;
                    xl_timeInRange = null;
                    xl_timeOutRange = null;
                    xl_totalHoursRange = null;

                    conAttRow++;

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
            Marshal.FinalReleaseComObject(newApplication);
        }

        /*

        //Converts/Populates the cells E5 to S5/T5 in shift code template sheet with appropriate dates based on shift period
        public void TemplateDateRangeEdit(Excel._Worksheet workSheet, string shiftPeriod, string shiftYear)
        {
            //int li_dayRangeCounter = 0;
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







        //This function processes the shift codes. Accepts ARRAY OF BADGE NUMBERS and ATTENDANCE PERIOD. Returns 1 if successful, 0 otherwise.
        public int ProcessShiftCode(string ID_NUMBER, DateTime DATE, DateTime DATE_FROM, DateTime DATE_TO)
        {
            Excel._Application newApplication;
            Excel._Workbook newBook, newBookTemplate;
            Excel._Worksheet newSheet, newSheetTemplate;

            newApplication = new Excel.Application();

            //Open the shift code template sheet
            try
            {
                newBookTemplate = newApplication.Workbooks.Open(lsTemplatePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                newSheetTemplate = (Excel._Worksheet)newBookTemplate.Worksheets.get_Item(1);
                newApplication.Visible = true;
            }
            catch
            {
                MessageBox.Show("ERROR: Shift code template cannot be found.");
                return 0;
            }


            newBook.Close(false, System.Type.Missing, System.Type.Missing);
            newSheetTemplate.Cells[3, 1] = "Attendance Period: " + shiftPeriod + ", " + (String.Format("{0:yyyy}", shiftYear));

            //Generate a filename for the processed shift code document
            //System prompts to save the generated shift codes worksheet.
            //If user cancels, generated shift code worksheet will not be saved, and system return 0 and terminates.             
            lsSaveDate = String.Format("{0:MMMddyyyy}", DateTime.Now);

            string saveFileName = cesection + " Shifts_" + shiftPeriod + "_" + lsSaveDate + ".xls";

            SaveFileDialog saveTemplateDialog = new SaveFileDialog();
            saveTemplateDialog.FileName = saveFileName;

            if (saveTemplateDialog.ShowDialog() == DialogResult.Cancel)
            {
                MessageBox.Show("Save cancelled. Terminating program.");
                return 0;
            }

            newBookTemplate.SaveAs(saveTemplateDialog.FileName, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

            //Close excel instances - save template, don't save shift code doc                                    
            newBookTemplate.Close(true, System.Type.Missing, System.Type.Missing);
            newApplication.Quit();

            //Close Excel, release objects, do garbage collection
            Marshal.ReleaseComObject(newSheetTemplate);
            Marshal.ReleaseComObject(newBookTemplate);
            Marshal.ReleaseComObject(newSheet);
            Marshal.ReleaseComObject(newBook);
            Marshal.ReleaseComObject(newApplication);

            newSheet = null;
            newSheetTemplate = null;
            newBook = null;
            newBookTemplate = null;
            newApplication = null;

            GC.WaitForPendingFinalizers();
            GC.Collect();

            //System return 1 if shift code generation is successful
            return 1;
        }




/*--------------------------------------------------------------------------------------------------------------------*/
/*--------------------------------------------------------------------------------------------------------------------*/
/*--------------------------------------------------------------------------------------------------------------------*/



        //Converts/Populates the cells E5 to S5/T5 in shift code template sheet with appropriate dates based on shift period
        /*
        public int TemplateDateRangeEdit(Excel._Worksheet workSheet, string shiftPeriod, string shiftYear)
        {
            int li_dayRangeCounter = 0;
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
                    li_dayRangeCounter++;
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
                    li_dayRangeCounter++;
                }
            }

            return li_dayRangeCounter;
        }
        */




        //Returns the shorthand month prefix used in naming attendance sheets
        public string GenerateFileMonth(string lsASFileName)
        {
            if (lsASFileName.IndexOf("Jan") != -1)
            { return "Jan"; }
            else if (lsASFileName.IndexOf("Feb") != -1)
            { return "Feb"; }
            else if (lsASFileName.IndexOf("Mar") != -1)
            { return "Mar"; }
            else if (lsASFileName.IndexOf("Apr") != -1)
            { return "Apr"; }
            else if (lsASFileName.IndexOf("May") != -1)
            { return "May"; }
            else if (lsASFileName.IndexOf("Jun") != -1)
            { return "June"; }
            else if (lsASFileName.IndexOf("Jul") != -1)
            { return "July"; }
            else if (lsASFileName.IndexOf("Aug") != -1)
            { return "Aug"; }
            else if (lsASFileName.IndexOf("Sep") != -1)
            { return "Sept"; }
            else if (lsASFileName.IndexOf("Oct") != -1)
            { return "Oct"; }
            else if (lsASFileName.IndexOf("Nov") != -1)
            { return "Nov"; }
            else if (lsASFileName.IndexOf("Dec") != -1)
            { return "Dec"; }
            else
            { return ""; }

        }


        //This function generates the shift code and posts it to the template      
        public string ConvertToShiftCode(DateTime refDatetime, DateTime timeIN, string WorkHours)
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

            else
            {
                if (timeIN == Convert.ToDateTime("01/01/0001 00:00:00")) //If TIME IN record is NULL or empty string, halt processing and return empty string
                {

                    return "E10";       /* ADDED ON FEBRUARY 24, 2012
                                         * Suggested by Charity. In case the engineer has no login info due to 
                                         * failure to swipe in, Logbox entry, or leave of absence, the system should generate
                                         * E10 as a default shift code.                                        
                                         */
                    //return "";
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

                        else if (dtTimeStamp >= Convert.ToDateTime("04:45 AM") && dtTimeStamp <= Convert.ToDateTime("06:45 AM"))
                        {
                            return "E81";
                        }

                        else if (dtTimeStamp >= Convert.ToDateTime("10:45 AM") && dtTimeStamp <= Convert.ToDateTime("12:45 PM"))
                        {
                            return "E82";
                        }

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
        public int colToInt(char column)
        {
            int colInt;
            colInt = Convert.ToInt32(column) - 65;
            return colInt;

        }






    }
}
