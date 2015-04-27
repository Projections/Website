using DotNet.Highcharts.Helpers;
using DotNet.Highcharts.Options;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;

namespace Projections_Capstone_Spring15
{


    public partial class Plot_Graph : System.Web.UI.Page
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        public static string excelFilepath = "";
        Object[] smoothedSSNList;
        Object[] monthlySSNList;
        Object[] altitudeList;
        string[] datesList;
        Object[] avgDoseInAllDataList;
        Computations c = new Computations();
        public static List<DateTime> StartdateList = new List<DateTime>();
        public static List<DateTime> EnddateList = new List<DateTime>();
        public static List<double> averagesList = new List<double>();

        protected void Page_Load(object sender, EventArgs e)
        {


        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            lblErrorDescription.Text = "";
            if (CheckCorrectExtension(btnTEPCBrowse))
            {
                try
                {
                    lblErrorDescription.Visible = false;
                    //imgTEPCLoading.Attributes.Add("style","display:block;");
                    excelFilepath = string.Concat(Server.MapPath(btnTEPCBrowse.FileName));
                    btnTEPCBrowse.PostedFile.SaveAs(excelFilepath);
                    c.CalculateAverage(excelFilepath);
                    WriteToExcelFile();
                    imgTEPCLoading.Attributes.Add("style", "display:none;");
                }
                catch (Exception ex)
                {
                    lblErrorDescription.Text = "Could not read the file";
                }
            }
            else
            {
                lblErrorDescription.Text = "File not recognized";
            }
        }


        private void WriteToExcelFile()
        {

            System.IO.StreamWriter sw = new System.IO.StreamWriter(string.Concat(Server.MapPath("AverageDoses.csv")));
            c.WriteToExcelFile(sw);
        }

        protected void lnkDownloadAvgTEPC_Click(object sender, EventArgs e)
        {
            Response.ContentType = "Application/x-msexcel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=AverageDoses.csv");
            Response.TransmitFile(Server.MapPath("AverageDoses.csv"));
            Response.End();
        }

        protected void btnUploadRAM_TLD_Click(object sender, EventArgs e)
        {
            if (CheckCorrectExtension(btnRAMBrowse))
            {
                try
                {
                    lblErrorDescription.Visible = false;
                    //imgTEPCLoading.Attributes.Add("style","display:block;");
                    excelFilepath = string.Concat(Server.MapPath(btnRAMBrowse.FileName));
                    btnRAMBrowse.PostedFile.SaveAs(excelFilepath);
                    // CreateExcelFile(excelFilepath);
                    // WriteToExcelFile();
                    if (datepickerStart.Text != null && datepickerEnd.Text != null)
                    {
                        double[] SMValues_Result = c.CalculateRAMValues(excelFilepath, datepickerStart.Text, datepickerEnd.Text);
                        for (int i = 1; i <= SMValues_Result.Length; i++)
                        {
                            lblSMValues.Text += "<br />Value of SM" + i + "is " + SMValues_Result[i - 1];
                        }
                    }
                    else
                    {

                        lblErrorDescription_RAM_TLD.Text = "Please Select Start and End Dates";
                    }
                    imgRAMLoading.Attributes.Add("style", "display:none;");
                }
                catch (Exception ex)
                {
                    lblErrorDescription.Text = "Could not read the file";
                }
            }
            else
            {
                lblErrorDescription.Text = "File not recongized";
            }
        }
        public bool CheckCorrectExtension(FileUpload uploadControl)
        {
            bool correctExtension = c.CheckCorrectExtension(uploadControl);
            return correctExtension;

        }

        protected void btnPlot_Click(object sender, EventArgs e)
        {
            DataTable dt = readFile();
            isValidData(dt);
            calculateAverages(dt);
            consolidatedData();
            //DotNet.Highcharts.Highcharts chart = new DotNet.Highcharts.Highcharts("chart").InitChart(new Chart { ZoomType = DotNet.Highcharts.Enums.ZoomTypes.X })
            //    .SetXAxis(new []{
            //    new XAxis
            //                {
            //                   Id="Axes1",
            //                    Categories = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" }
            //                },
            //                new XAxis
            //                {
            //                     Opposite=true,
            //                     Id="Axes2",
            //                    Categories=new[]{"Sun", "Mon","Tue","Wed","Thu","Fri","Sat"}
            //                },
            //                new XAxis
            //                { 
            //                     Id="Axes3",
            //                    Categories=new[]{"Summer", "Winter","Fall"}
            //                }
            //    })
            //    //.SetYAxis(new YAxis
            //    //{
            //    //    Categories = new[] { "0", "50", "100", "150", "200", "250", "300" }
            //    //}
            //    //)
            //    .SetSeries(new[]
            //    { new Series
            //                {
            //                    XAxis="Axes1",
            //                    Name="First Series",
            //                    Data = new Data(new object[] { 29.9, 71.5, 106.4, 129.2, 144.0, 176.0, 135.6, 148.5, 216.4, 194.1, 95.6, 54.4 })
            //                },
            //        new Series
            //                {
            //                    XAxis="Axes2",
            //                    Name="Second series",
            //                    Data = new Data(new object[] { 129.9, 171.5, 106.4, 129.2, 144.0, 176.0, 35.6 })
            //                },
            //    new Series
            //                {
            //                    XAxis="Axes3",
            //                    Name="Second series",
            //                    Data = new Data(new object[] { 100, 120, 95 })
            //                }
            //    });
            DotNet.Highcharts.Highcharts chart = new DotNet.Highcharts.Highcharts("chart").InitChart(new Chart
            {
                ZoomType = DotNet.Highcharts.Enums.ZoomTypes.X,
            })
            .SetXAxis(new[]{
                new XAxis
                            {
                             Id="Month Axes",
                              // Type=DotNet.Highcharts.Enums.AxisTypes.Datetime,
                                Categories = datesList,
                               Labels=new XAxisLabels{Step=10, StaggerLines=1}
                             // MinRange=30*24
                            }
                           
                , new XAxis 
                            { 
                                Id="new X",
                                Categories=new [] {"Category1", "Category2", "Category3"}
                            }
            });
            chart.SetTitle(new Title { Text = "Space Weather and Altitude" });
            chart.SetSeries(new[]
                { new Series
                            {
                                
                                YAxis="Sunspot",
                                XAxis="Month Axes",
                                Name="Smoothed SSN",
                                Data = new Data(smoothedSSNList)
                                //PlotOptionsLine=new PlotOptionsLine{PointInterval=24*24*3600000, PointStart=new PointStart(Convert.ToDateTime(datesList[0]))}
                            },
                    new Series
                            {
                                XAxis="Month Axes",
                                YAxis="Altitude",   
                                Name="Altitude",
                                Data = new Data(altitudeList)
                            },
                    new Series
                            {
                                XAxis="Month Axes",
                                YAxis="Sunspot",
                                Name="Monthly SSN",
                                Data = new Data(monthlySSNList)
                            } ,
                    new Series
                            {
                                XAxis="Month Axes",
                                 YAxis="Altitude",   
                                Name="Average Dose Values",
                                Data = new Data(avgDoseInAllDataList)
                            } ,
                    new Series
                            {
                            YAxis="Altitude",
                            //XAxis="new X",
                            Type=DotNet.Highcharts.Enums.ChartTypes.Columnrange,
                            Name="Dummy Data",
                           // Data=new Data(new object[,]{{low=12.9,high=6.8},{12.5,99.8},{32.5,56.4}})
}
                });
            chart.SetYAxis(new[]{
                   new YAxis
                   {
                       Id="Sunspot",
                       Min=0,
                       Max=400,
                       TickInterval=25,
                       Title=new YAxisTitle { Text = "Sunspot Number" }//,
                       //Labels=new YAxisLabels{Format="{value} km"}
                   },
                   new YAxis
                   {
                       Id="Altitude",
                       Min=0,
                       Max=500,
                       TickInterval=30,
                       Opposite=true,
                       Title=new YAxisTitle { Text = "Altitude [km] and dose values [µGy]" }
                   }}
                );
            ltrChart.Text = chart.ToHtmlString();

        }

        public void consolidatedData()
        {
            string alldatafilePath = Server.MapPath("DataTillDate.xlsx");
            DataTable allData = getDataTable(alldatafilePath);
            var data = allData;
            int rowCount = allData.Rows.Count;
            avgDoseInAllDataList = new Object[rowCount];
            altitudeList = new Object[rowCount];
            monthlySSNList = new Object[rowCount];
            smoothedSSNList = new Object[rowCount];
            datesList = new string[rowCount];
            int countForeach = 0;
            foreach (DataRow dR in allData.Rows)
            {
                try
                {
                    DateTime day = DateTime.Parse(dR[0].ToString());
                    datesList[countForeach] = day.Month + "/" + day.Year;
                    avgDoseInAllDataList[countForeach] = dR[1];
                    altitudeList[countForeach] = dR[2];
                    monthlySSNList[countForeach] = dR[3];
                    smoothedSSNList[countForeach] = dR[4];
                    countForeach++;
                }
                catch (Exception exc)
                {

                }
            }

        }

        public static DataTable getDataTable(string path)
        {
            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String[] s = getsheet(path);
            for (int i = 0; i < s.Length; i++)
            {
                String query = "SELECT * FROM [" + s[0] + "]"; // You can use any different queries to get the data from the excel sheet
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
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties=Excel 12.0 xml;";
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

        public static DataTable readFile()
        {
            DataTable firstTable = new DataTable();
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;
            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            String[] excelSheets;
            string excelFile = "Target CPD.xlsx";
            try
            {
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties=Excel 12.0 xml;";
                objConn = new OleDbConnection(connString);
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    // return null;
                    Console.WriteLine("No data found");
                }
                excelSheets = new String[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
                String query = "SELECT MissionStart,MIssionStop,ProxyCPDData FROM [" + excelSheets[0] + "]"; // You can use any different queries to get the data from the excel sheet
                OleDbConnection conn = new OleDbConnection(connString);

                cmd = new OleDbCommand(query, conn);
                da = new OleDbDataAdapter(cmd);
                da.Fill(ds);
                firstTable = ds.Tables[0];
                // return firstTable;
                return firstTable;
            }
            catch (Exception exc)
            {

            }

            return firstTable;
        }

        public static bool isValidData(DataTable dt)
        {
            // bool isDataValid = true;
            //int count = 0;
            for (int i = 0; i < dt.Rows.Count - 1; i++)
            {
                DataRow dr1 = dt.Rows[i];
                DataRow dr2 = dt.Rows[i + 1];
                DateTime endDay = DateTime.Parse(dr1[1].ToString());
                DateTime nextBeginDay = DateTime.Parse(dr2[0].ToString());
                double dayDifference = endDay.Subtract(nextBeginDay).TotalDays;
                if (dayDifference < 0.0)
                {
                    Console.WriteLine("Error in line: " + i);
                    Console.ReadKey();
                    return false;
                }
            }
            //Console.WriteLine("Read Successful");
            // Console.ReadKey();
            return true;
        }

        public static void calculateAverages(DataTable dt)
        {
            DateTime StartDay, EndDay;
            foreach (DataRow dr in dt.Rows)
            {
                //GEt the first day and last day
                StartDay = Convert.ToDateTime(dr[0].ToString());
                EndDay = Convert.ToDateTime(dr[1].ToString());

                //Get the total number of days
                double TotalDays = Math.Ceiling((EndDay.Date - StartDay.Date).TotalDays + 1);

                //Calculate the average CPD for the given period
                double AvgCPD = Convert.ToDouble(dr[2]) / TotalDays;

                //Add the average CPD, first month and last month values to the list
                averagesList.Add(AvgCPD);
                StartdateList.Add(new DateTime(StartDay.Year, StartDay.Month, 1));
                EnddateList.Add(new DateTime(EndDay.Year, EndDay.Month, 1));


            }
        }

    }
}