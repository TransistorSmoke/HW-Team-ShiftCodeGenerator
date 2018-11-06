using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

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
        }

        public int ValidateArguments(string year, bool hoursCheck, string dateFrom_8hours, string dateTo_8hours)
        {
            string checkYear = "", 
                   checkDateFrom = "", 
                   checkDateTo = "", 
                   checkHoursCheck = "";
            int check8Hours = 1;
            string valMessage = "ERRORS FOUND:\n-----------------\n\n";

            if (year == "")
            {
                checkYear = null;  
                valMessage = valMessage + "YEAR must not be empty.\n";                
            }

            if(hoursCheck = true && (dateFrom_8hours == "" || dateTo_8hours == ""))
            {
                check8Hours = 0;
                valMessage = valMessage + "8-hour shift period checked. Please supply appropriate DATE RANGE.\n";
            }

            if(checkYear == null || checkDateFrom  == null || checkDateTo == null || check8Hours == 0)
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
            //ValidateArguments(ddShiftYear.Text, radioButtonYes.Checked, ddFrom.Text, ddTo.Text);
            
            
            string year = ddShiftYear.Text;
            string shiftPeriod = ddShiftPeriod.Text;
            string shiftMonth = shiftPeriod.Substring(0, 3);
            string shiftDayStart = shiftPeriod.Substring(4, 2);
            string shiftDayEnd = shiftPeriod.Substring(7,2);

            string month = MONTH_STRING_TO_INT(shiftMonth);

            string dateFROM = year + "/" + month + "/" + shiftDayStart;
            string dateTO = year + "/" + month + "/" + shiftDayEnd;
                                                                       
            //Connect to MySQL database
            
            MySqlConnection conn = new MySqlConnection();
            string connString = "server=localhost;uid=root;pwd=;database=db_shiftcode;";

            try
            {
                conn.ConnectionString = connString;
                conn.Open();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                conn.Close();
            }


            string query = "SELECT * FROM m_attendance WHERE date BETWEEN '" + dateFROM + "' AND '" + dateTO + "';";
            

            MySqlCommand sqlQuery = new MySqlCommand(query, conn);
            MySqlDataReader sqlReader = sqlQuery.ExecuteReader();

            /* Summary of Query Results
             * ------------------------
             * GetString(0) ===> badgeno
             * GetString(2) ===> name
             * GetString(3) ===> date
             * GetString(4) ===> time_in
             * GetString(5) ===> time_out
             * GetString(6) ===> total_hours
             */

            while (sqlReader.Read())
            {                                                                                           
                MessageBox.Show("BADGE NUMBER\t: " + sqlReader.GetString(0) + "\n" +
                                "FULL NAME\t: " + sqlReader.GetString(1) + "\n" +
                                "DATE\t\t: " + sqlReader.GetString(2) + "\n" +
                                "TIME IN\t\t: " + sqlReader.GetString(3) + "\n" +
                                "TIME OUT\t\t: " + sqlReader.GetString(4) + "\n" +
                                "TOTAL HOURS\t: " + sqlReader.GetString(5));            
            }
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








    }
}
