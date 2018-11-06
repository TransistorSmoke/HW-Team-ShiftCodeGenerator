using System;
using System.IO;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace HWAttendanceGrabber
{
    class configFunction
    {
        public string getConfigFilePath()
        {
            string configPath = AppDomain.CurrentDomain.BaseDirectory + "config.ini";
            return configPath;
        }

        public string getBadgeFilePath()
        {
            string badgePath = AppDomain.CurrentDomain.BaseDirectory + "badge.ini";
            return badgePath;
        }
       

        //public string getAttendanceFilePath(string year) - NOT USED ANYMORE
        public string getAttendanceFilePath()
        {
            configFile ConfigManager = new configFile(getConfigFilePath());

            string attendanceFilePath = ConfigManager.Read("PATH", "attendancePath");

            /*
            DirectoryInfo attendanceDirectory = new DirectoryInfo(attendanceFilePath);
            string[] ls_ASMonthSheetArray = Directory.GetFiles(attendanceFilePath, "*.xls");
            
            
            foreach (DirectoryInfo folder in attendanceDirectory.GetDirectories())
            {
                if (folder.Name.Contains(year))
                {
                    attendanceFilePath = attendanceFilePath + folder.Name;
                }

            }
             */

            return attendanceFilePath;
        }



        public int SaveToDB(string year)
        {
            return 1;
        }

        public ArrayList getBadgeArray(string badgePath, string section)
        {
            string id;
            int badgeCount = 0;
            ArrayList arrayBadge = new ArrayList();

            configFile newConfig = new configFile(getBadgeFilePath());

            //RETRIEVE REGISTERED BADGE NUMBERS FROM CONFIG.INI
            do
            {
                id = newConfig.Read(section, "idNum[" + badgeCount + "]");
                arrayBadge.Add(id);
                badgeCount++;
            }
            while (id != "");
            //END OF ID RETRIEVAL
            return arrayBadge;
       
        }



    }
}
