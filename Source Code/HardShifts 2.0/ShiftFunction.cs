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
        public string getINIPath()
        {
            string INIPath = AppDomain.CurrentDomain.BaseDirectory + "config.ini";
            return INIPath;
        }

        public string getINIBadgePath()
        {
            string INIBadgePath = AppDomain.CurrentDomain.BaseDirectory + "badge.ini";
            return INIBadgePath;
        }


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


        public string getTemplatePath(string cesection)
        {
            string templateSheetPath;
            INIFile newTemplateINI = new INIFile(getINIPath());

            templateSheetPath = newTemplateINI.Read("PATH", "templatePath");

            if (cesection == "HARDWARE DEVELOPMENT")
            {
                templateSheetPath = AppDomain.CurrentDomain.BaseDirectory + templateSheetPath + "HW_template.xlsx";
            }
            else if (cesection == "SOFTWARE DEVELOPMENT")
            {
                templateSheetPath = AppDomain.CurrentDomain.BaseDirectory + templateSheetPath + "SWD_template.xlsx";
            }
            else if (cesection == "SOFTWARE VALIDATION")
            {
                templateSheetPath = AppDomain.CurrentDomain.BaseDirectory + templateSheetPath + "SWV_template.xlsx";
            }
            else if (cesection == "ELECTROMECHANICAL")
            {
                templateSheetPath = AppDomain.CurrentDomain.BaseDirectory + templateSheetPath + "EM_template.xlsx";
            }
            else if (cesection == "LAB, PM AND PQ")
            {
                templateSheetPath = AppDomain.CurrentDomain.BaseDirectory + templateSheetPath + "LAB_template.xlsx";
            }
            else if (cesection == "ALL")
            {
                templateSheetPath = AppDomain.CurrentDomain.BaseDirectory + templateSheetPath + "ALL_template.xlsx";
            }

            return templateSheetPath;
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
