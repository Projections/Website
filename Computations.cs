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
            sw.WriteLine("Date"+","+ "Average_Doses");
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
            Excel.Range range;
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(getExcelFilePath); //Giving the path to excel workbook to open and read the file.
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here

            string str = "";
            int rCnt, cCnt, dateCol = 0, doseCol = 0;
            range = MySheet.UsedRange; //range contains all the cells taht are currently in use.
            decimal avgOfDose = 0;
            int count = 0;
            //loop to iterate through the columns which contain Total and  Date, which we use to find the average values!
            //Cursor points to the required fields.
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                str = Convert.ToString((range.Cells[1, cCnt] as Excel.Range).Value2);
                if (!String.IsNullOrEmpty(str))
                {
                    if (str.Contains("Date")) // Storing the column number in dateCol which has Date in the first row
                    {
                        dateCol = cCnt;

                    }
                    if (str.Contains("Total")) // Storing the column number in doseCol which has Total in the first row.
                    {

                        doseCol = cCnt;
                    }
                }
            }
            //loop to iterate through the rows that are taken from the first loop, this takes all the date values and total dose values in the selected file. 
            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {

                DateTime dateOfTEPC = new DateTime();
                decimal doseValue = 0;
                bool isValidTEPCDate = false; // Boolean to check for valid TEPC date
                bool isValidDoseValue = false; // Boolean to check for valid Dose Value


                object inpDate = (range.Cells[rCnt, dateCol] as Excel.Range).Value2; // gets the range of cells, rCnt is the rown count which contains the current row number of the particular column number of date
                if (inpDate != null)
                {
                    if (inpDate is double) // check if date is double
                    {
                        dateOfTEPC = DateTime.FromOADate((double)inpDate); // conver from OADate to Date data type
                        isValidTEPCDate = true;
                    }
                    else
                    {
                        DateTime.TryParse((string)inpDate, out dateOfTEPC); //  Parsing the input date to DateTime
                        isValidTEPCDate = true;
                    }


                }


                object dValue = (range.Cells[rCnt, doseCol] as Excel.Range).Value2; // Gets the range of cells which contain dose values, rCnt is the variable which iterated through all the dose values  of the aprticular row.
                if (dValue != null) //check for dose value for null. the cell in the excel may be empty
                {
                    if (dValue is double)
                    {
                        doseValue = Convert.ToDecimal(dValue);
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
                        // listOfYearValues.Add(curYear);
                        listOfAvgDoseValues.Add(avgOfDose);
                        count = 1; //Updating the count value for further iterations of for loop
                        //  MessageBox.Show(avgOfDose.ToString());
                        avgOfDose = doseValue;
                    }
                }

            }
            listOfAvgDoseValues.Add(avgOfDose / count); // Finally adding the average dose values to the list
            try
            {
                MyBook.Close(true, getExcelFilePath, null);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {

            }
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
            //Console.WriteLine("Dates should be between 2000-12-01 and 2013-11-07");
            //Console.WriteLine("Enter the start date:(yyyy-mm-dd)");//User enters the start and end dates
            //string start = Console.ReadLine();
            //    Console.WriteLine("Enter The end date:(yyyy-mm-dd)");
           // string end = Console.ReadLine();
            //string startdate = DateTime.ParseExact(start,
            //                       "yyyy-MM-dd",
            //                        CultureInfo.InvariantCulture).ToString();
            //string enddate = DateTime.ParseExact(end,
            //                       "yyyy-MM-dd",
            //                        CultureInfo.InvariantCulture).ToString(); ;
            double o = (DateTime.Parse(enddate) - DateTime.Parse(startdate)).TotalDays;//Difference of start and end dates
            DataTable d = s(path);//XLS file is taken as a data table
            if (o > 185 || o < 31)
            {
                  // "The days between start and end dates should be betwee 31 and 185 days");
                //Environment.Exit(0);
            }
            var d1 = d.Rows[0];//Considering the required rows from the data table
            DataRow[] result = d.Select("enddate >= '" + startdate + "' AND startdate<='" + enddate + "' ");
            result = result.Distinct().ToArray();//An array list is taken for start and end dates
            var foos = new HashSet<DataRow>(result);
            int k = 0, j = result.Length - 1;
            HashSet<int> p = new HashSet<int>();


            result = foos.ToArray();
            double[] sum; // declare numbers as an int array of any size
            sum = new double[4];


            double numofdays = 0;//Initialize the number of days is zero
            for (int i1 = 0; i1 < result.Length; i1++)
            {
                var number = result[i1][3];//Here considering the expodays and absored dose values from the XLS file
                var dose = result[i1][5];
                DateTime start1;
                string loc = result[i1][2].ToString();
                DateTime s1 = DateTime.Parse(result[i1][0].ToString());
                DateTime s2 = DateTime.Parse(result[i1][1].ToString());
                float wt1 = float.Parse(dose.ToString()) / float.Parse(number.ToString());//Calculating the wt1=absored days/expodays
                int i = DateTime.Compare(DateTime.Parse(result[i1][1].ToString()), DateTime.Parse(enddate));
                if (i < 0)//Considering the user enter dates with each location from the XLS file and calculating the wt1 for 4 locations
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

            //Finally calculating the required dose value for given dates

            //foreach (double d12 in sum)
            // Console.WriteLine(d12);
            //Console.WriteLine("The dose value in SM-1 is: " + sum[0]);//Displaying the dose values for each location
            //Console.WriteLine("The dose value in SM-2 is: " + sum[1]);
            //Console.WriteLine("The dose value in SM-3 is: " + sum[2]);
            //Console.WriteLine("The dose value in SM-4 is: " + sum[3]);

            //Console.ReadKey();

        }
        public static DataTable s(string path)
        {

            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String[] s = getsheet(path);
            for (int i = 0; i < s.Length; i++)
            {
                String query = "SELECT * FROM ["+s[0]+"]"; // You can use any different queries to get the data from the excel sheet
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
                    // Exception Msg 

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

                String connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                  "Data Source=" + excelFile + ";Extended Properties=Excel 8.0;";

               // String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            
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