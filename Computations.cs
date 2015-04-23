using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data.OleDb;
using System.Web;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace Projections_Capstone_Spring15
{
    public class Computations
    {
        public static int curMonth = 0;
        public static int curYear = 0;
        public static bool initDateFlag = true;
        List<decimal> listOfAvgDoseValues = new List<decimal>();
        List<string> listOfdateValues = new List<string>();
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        public void WriteToExcelFile(System.IO.StreamWriter sw)
        {
            sw.WriteLine("Date" + "," + "Average_Doses");
            for (int ii = 0; ii < listOfAvgDoseValues.Count; ii++)
            {
                sw.WriteLine(listOfdateValues[ii] + " ,  " + listOfAvgDoseValues[ii]); // Writing the obtained Avg Dose values into a CSV file with respect to the corresponding dates.
            }
            sw.Close();
        }
        public bool CheckCorrectExtension(FileUpload uploadControl)
        {
            bool correctExtension = false;
            if (uploadControl.HasFile)
            {
                string fileExtension = Path.GetExtension(uploadControl.FileName).ToLower();
                string[] extensionsAllowed = { ".xls", ".xlsx", ".csv", ".xlsm" };

                for (int i = 0; i < extensionsAllowed.Length; i++)
                {
                    if (fileExtension == extensionsAllowed[i])
                    {
                        correctExtension = true;
                    }
                }
            }
            return correctExtension;
        }


        public void CalculateAverage(string getExcelFilePath)
        {
            DataTable forAvgDoseValues = s(getExcelFilePath, "Date, Total");
            decimal avgOfDose = 0;
            int count = 0;
            foreach (DataRow dr in forAvgDoseValues.Rows)
            {
                DateTime dateOfTEPC = new DateTime();
                decimal doseValue = 0;
                bool isValidTEPCDate = false; // Boolean to check for valid TEPC date
                bool isValidDoseValue = false; // Boolean to check for valid Dose Value
                //object inpDate = dr[0].ToString().Trim();
                if (!string.IsNullOrEmpty(dr[0].ToString()))
                {
                    try
                    {
                        if (dr[0] is double) // check if date is double
                        {
                            dateOfTEPC = DateTime.FromOADate((double)dr[0]); // conver from OADate to Date data type
                            isValidTEPCDate = true;
                        }
                        else
                        {
                            dateOfTEPC = DateTime.Parse(dr[0].ToString());
                             //  Parsing the input date to DateTime
                            isValidTEPCDate = true;
                        }
                    }
                    catch(Exception exc)
                    {

                    }
                }
                object inpDose = dr[1];
                if (inpDose != null) //check for dose value for null. the cell in the excel may be empty
                {
                    if (inpDose is double)
                    {
                        doseValue = Convert.ToDecimal(inpDose);
                        isValidDoseValue = true;
                    }
                }
                if (isValidDoseValue && isValidTEPCDate) // Check if both Dose and TEPCDate are valid
                {
                    if (initDateFlag) // Initial date flag to check for the cursor reaches another month row.
                    {
                        curMonth = dateOfTEPC.Month;
                        curYear = dateOfTEPC.Year;
                        listOfdateValues.Add(dateOfTEPC.ToOADate().ToString()); // Adding corresponding date to listOdateValues list.
                        initDateFlag = false; // Update the boolean 
                    }
                    if (dateOfTEPC.Month == curMonth && dateOfTEPC.Year == curYear) // Incrementing the count (used to divide sum of dose values with) 
                    //and adding the dose values by checking current month and year
                    {
                        count++;
                        avgOfDose += doseValue;
                    }
                    else
                    {
                        curMonth = dateOfTEPC.Month;
                        curYear = dateOfTEPC.Year;
                        avgOfDose = avgOfDose / count; // calculating the average dose value
                        listOfdateValues.Add(dateOfTEPC.ToOADate().ToString());
                        listOfAvgDoseValues.Add(avgOfDose);
                        count = 1; //Updating the count value for further iterations of for loop
                        avgOfDose = doseValue;
                    }
                }
            }
            // Finally adding the average dose values to the list
            listOfAvgDoseValues.Add(avgOfDose / count);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
        }

        public double[] CalculateRAMValues(string getExcelFilePath, string startdate, string enddate)
        {
            string path = getExcelFilePath;
            double o = (DateTime.Parse(enddate) - DateTime.Parse(startdate)).TotalDays;//Difference of start and end dates
            DataTable d = s(path, "*");//XLS file is taken as a data table
            if (o > 185 || o < 31)
            {
                // "The days between start and end dates should be betwee 31 and 185 days");
                //Environment.Exit(0);
            }
            var d1 = d.Rows[0];//Considering the required rows from the data table
            DataRow[] result = d.Select("enddate >= '" + startdate + "' AND startdate<='" + enddate + "' ");
            result = result.Distinct().ToArray();//An array list is taken for start and end dates
            var foos = new HashSet<DataRow>(result);
            int j = result.Length - 1;
            HashSet<int> p = new HashSet<int>();
            result = foos.ToArray();
            double[] sum; // declare numbers as an int array of any size
            sum = new double[4];
            double numofdays = 0;//Initialize the number of days is zero
            for (int i1 = 0; i1 < result.Length; i1++)
            {
                var number = result[i1][3];//Here considering the expodays and absorbed dose values from the XLS file
                var dose = result[i1][5];
                DateTime start1;
                string loc = result[i1][2].ToString();
                DateTime s1 = DateTime.Parse(result[i1][0].ToString());
                DateTime s2 = DateTime.Parse(result[i1][1].ToString());
                float wt1 = float.Parse(dose.ToString()) / float.Parse(number.ToString());//Calculating the wt1=absored days/expodays
                int i = DateTime.Compare(DateTime.Parse(result[i1][1].ToString()), DateTime.Parse(enddate));
                if (i < 0)//Considering the user input dates with each location from the XLS file and calculating the wt1 for 4 locations
                {
                    j = DateTime.Compare(DateTime.Parse(startdate), s1);
                    if (j == 0)
                    {
                        start1 = s1;
                    }
                    else if (j < 0)
                    {
                        start1 = s1;
                    }
                    else
                    {
                        start1 = DateTime.Parse(startdate);
                    }
                    numofdays = (start1 - s2).TotalDays;
                }
                else if (i > 0)
                {
                    j = DateTime.Compare(DateTime.Parse(startdate), s1);
                    if (j == 0)
                    {
                        start1 = s1;
                    }
                    else if (j < 0)
                    {
                        start1 = s1;
                    }
                    else
                    {
                        start1 = DateTime.Parse(startdate);
                    }
                    numofdays = (start1 - DateTime.Parse(enddate)).TotalDays;
                }
                else
                {
                    j = DateTime.Compare(DateTime.Parse(startdate), s1);
                    if (j == 0)
                    {
                        start1 = s1;
                    }
                    else if (j < 0)
                    {
                        start1 = s1;
                    }
                    else
                    {
                        start1 = DateTime.Parse(startdate);

                    }
                    numofdays = (start1 - s2).TotalDays;
                }
                if (loc == "(SM-1)")
                    sum[0] += (wt1 * (-numofdays));
                if (loc == "(SM-2)")
                    sum[1] += (wt1 * (-numofdays));
                if (loc == "(SM-3)")
                    sum[2] += (wt1 * (-numofdays));
                if (loc == "(SM-4)")
                    sum[3] += (wt1 * (-numofdays));
            }
            return sum;
        }
        public static DataTable s(string path, string columnNames)
        {
            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String[] s = getsheet(path);
            for (int i = 0; i < s.Length; i++)
            {
                String query = "SELECT " + columnNames + " FROM [" + s[0] + "]"; // You can use any different queries to get the data from the excel sheet
                OleDbConnection conn = new OleDbConnection(connString);
                if (conn.State == ConnectionState.Closed) conn.Open();
                try
                {
                    cmd = new OleDbCommand(query, conn);
                    da = new OleDbDataAdapter(cmd);
                    da.Fill(ds);
                    DataTable firstTable = ds.Tables[0];
                    return firstTable;
                }
                catch
                {
                     return null;
                }
                finally
                {
                    da.Dispose();
                    conn.Close();
                }
            }
            return null;
        }
        private static String[] getsheet(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;
            try
            {
               String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                objConn = new OleDbConnection(connString);
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    return null;
                }
                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
                return excelSheets;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

    }
}