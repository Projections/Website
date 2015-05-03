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
using System.Text;
using System.Threading.Tasks;

namespace Projections_Capstone_Spring15
{


    public partial class Plot_Graph : System.Web.UI.Page
    {
        #region Public Variables
        public static string excelFilepath = "";
        Object[] smoothedSSNList;
        Object[] monthlySSNList;
        Object[] altitudeList;
        string[] datesList;
        Object[] avgDoseInAllDataList;

        string[] strEndDate;
        string[] strStartDate;

        dynamic[] loc1;
        dynamic[] loc2;
        dynamic[] loc3;
        dynamic[] loc4;
        static dynamic[] cpd;
        //static dynamic[] cpd2;
        Computations c = new Computations();

        #endregion

          #region Page Load
        protected void Page_Load(object sender, EventArgs e)
        {


        }
          #endregion

        #region Upload TEPC file
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
        #endregion


        #region write to Excel File

        private void WriteToExcelFile()
        {

            System.IO.StreamWriter sw = new System.IO.StreamWriter(string.Concat(Server.MapPath("AverageDoses.csv")));
            c.WriteToExcelFile(sw);
        }
        #endregion

        #region Download Average Dose Values Link
        protected void lnkDownloadAvgTEPC_Click(object sender, EventArgs e)
        {
            Response.ContentType = "Application/x-msexcel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=AverageDoses.csv");
            Response.TransmitFile(Server.MapPath("AverageDoses.csv"));
            Response.End();
        }
        #endregion

        #region Upload RAM Data File
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

        #endregion

        #region check for extensions of uploaded file
        public bool CheckCorrectExtension(FileUpload uploadControl)
        {
            bool correctExtension = c.CheckCorrectExtension(uploadControl);
            return correctExtension;

        }
        #endregion

        #region Plot Button Click
        protected void btnPlot_Click(object sender, EventArgs e)
        {
            RAMLocWiseValues();
            //CPD Data
            DataTable dt = readFile();
            isValidData(dt);
            calculateAverages(dt);
            consolidatedData();
            //Plot the first chart with cosolidated data
            plotConsolidatedData();
            //plot the second chart with RAM and CPD data
            plotRAMData();
        }
        #endregion

        #region RAM Data Plot
        private void plotRAMData()
        {
            //RAM graph
            DotNet.Highcharts.Highcharts RAMChart = new DotNet.Highcharts.Highcharts("chart1").InitChart(new Chart
            {
                ZoomType = DotNet.Highcharts.Enums.ZoomTypes.Y,
                Type = DotNet.Highcharts.Enums.ChartTypes.Columnrange,
                Inverted = true
            })
            .SetXAxis(new[]{
                new XAxis
                            {
                             Type=DotNet.Highcharts.Enums.AxisTypes.Linear,
                             Max=350,
                             Reversed=false,
                             Title=new XAxisTitle{Text="Dose values"}
                            }
        }).SetYAxis(new[]{
        new YAxis{
             Type=DotNet.Highcharts.Enums.AxisTypes.Datetime,
             Title=new YAxisTitle{Text="Date"}
            }
        });
            RAMChart.SetTitle( new Title{Text="RAM and CPD"})
            .SetTooltip(new Tooltip
            {
                PointFormat = "{point.low:%e %b, %y} - {point.high:%e %b, %y}",
                HeaderFormat = "<b>{series.name}:</b>{point.x}<br />"
            });
            RAMChart.SetSeries(new[]
                {  new Series
                            {
                                Name="CPD",
                                Data = new Data(cpd),
                                Color= System.Drawing.Color.Red
                            },
                    new Series
                            {
                                Name="SM-1",
                                Data = new Data(loc1),
                            },
                             new Series
                            {
                                Name="SM-2",
                                Data = new Data(loc2)
                            },
                             new Series
                            {
                                Name="SM-3",
                                Data = new Data(loc3)
                            },
                             new Series
                            {
                                Name="SM-4",
                                Data = new Data(loc4)
                            }});
            ltrRAM.Text = RAMChart.ToHtmlString();
        }
        #endregion

        #region All Data Plot

        private void plotConsolidatedData()
        {
            //All data graph
            DotNet.Highcharts.Highcharts chart = new DotNet.Highcharts.Highcharts("chart").InitChart(new Chart
            {
                ZoomType = DotNet.Highcharts.Enums.ZoomTypes.X,

            })
            .SetXAxis(new[]{
                new XAxis
                            {
                             Id="Month Axes",
                                Categories = datesList,
                               Labels=new XAxisLabels{Step=15, StaggerLines=1}
                             // MinRange=30*24
                            }
                //, new XAxis 
                //            { 
                //                Id="RAM_X", 
                //                Type=DotNet.Highcharts.Enums.AxisTypes.Datetime,
                //            // Max=350
                //            }
            });

            chart.SetTitle(new Title { Text = "Space Weather and Altitude" });
            chart.SetSeries(new[]
                { new Series
                            {
                                YAxis="Sunspot",
                                XAxis="Month Axes",
                                Name="Smoothed SSN",
                                Data = new Data(smoothedSSNList)
                               // PlotOptionsLine=new PlotOptionsLine{PointInterval=24*24*3600000, PointStart=new PointStart(Convert.ToDateTime(datesList[0]))}
                            },
//                            new Series
//                            {
//                                Name="CPD",
//                                Data = new Data(cpd2),
//                                Color= System.Drawing.Color.Red,
//                                Type=DotNet.Highcharts.Enums.ChartTypes.Columnrange,
//                                XAxis="RAM_X",
//                                YAxis="RAM_Y",
//                                PlotOptionsColumnrange=new PlotOptionsColumnrange{
//                                Tooltip=new PlotOptionsColumnrangeTooltip{
//                                    PointFormat = "{point.low:%e %b, %y} - {point.high:%e %b, %y}",
//                HeaderFormat = "<b>{series.name}:</b>{point.y}<br />"
//}
//                                }
                //                {PointFormat = "{point.low:%e %b, %y} - {point.high:%e %b, %y}",
                //HeaderFormat = "<b>{series.name}:</b>{point.x}<br />"})
                          //  },
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
                            }
                });
            chart.SetYAxis(new[]{
                   new YAxis
                   {
                       Id="Sunspot",
                       Min=0,
                       Max=400,
                       TickInterval=25,
                       Title=new YAxisTitle { Text = "Sunspot Number" }
                   },
                   new YAxis
                   {
                       Id="Altitude",
                       Min=0,
                       Max=500,
                       TickInterval=30,
                       Opposite=true,
                       Title=new YAxisTitle { Text = "Altitude [km] and dose values [µGy]" }
                   },
            //         new YAxis{
            // Type=DotNet.Highcharts.Enums.AxisTypes.Linear,
            // Id="RAM_Y",
             
            //}
            });
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
        #endregion

        #region Table read Logic
        public DataTable getDataTable(string path)
        {
            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String[] s = getsheet(path);
            for (int i = 0; i < s.Length; i++)
            {
                String query = "SELECT * FROM [" + s[i] + "]"; // You can use any different queries to get the data from the excel sheet
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
        private String[] getsheet(string excelFile)
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
        public DataTable readFile()
        {
            DataTable firstTable = new DataTable();
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;
            OleDbCommand cmd = new OleDbCommand();//This is the OleDB data base connection to the XLS file
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            String[] excelSheets;
            string excelFile = Server.MapPath("Target CPD.xlsx");
            try
            {
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties=Excel 12.0 xml;";
                objConn = new OleDbConnection(connString);
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    // return null;
                   // Console.WriteLine("No data found");
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
                return firstTable;
            }
            catch (Exception exc)
            {

            }

            return firstTable;
        }
        #endregion

        #region check for valid Data
        public static bool isValidData(DataTable dt)
        {
           for (int i = 0; i < dt.Rows.Count - 1; i++)
            {
                DataRow dr1 = dt.Rows[i];
                DataRow dr2 = dt.Rows[i + 1];
                DateTime endDay = DateTime.Parse(dr1[1].ToString());
                DateTime nextBeginDay = DateTime.Parse(dr2[0].ToString());
                double dayDifference = endDay.Subtract(nextBeginDay).TotalDays;
                if (dayDifference < 0.0) //Check if the data is continuous and return false if there is data missing
                {
                    //Console.WriteLine("Error in line: " + i);
                    //Console.ReadKey();
                    return false;
                }
            }
            return true;
        }
        #endregion

        #region logic to calculate AVG Values of Dose
        public static void calculateAverages(DataTable dt)
        {
            DateTime StartDay, EndDay;
            cpd = new dynamic[dt.Rows.Count];
            //cpd2 = new dynamic[dt.Rows.Count];
            int index = 0;
            foreach (DataRow dr in dt.Rows)
            {
                //GEt the first day and last day
                StartDay = Convert.ToDateTime(dr[0].ToString());
                EndDay = Convert.ToDateTime(dr[1].ToString());

                //Get the total number of days
                double TotalDays = Math.Ceiling((EndDay.Date - StartDay.Date).TotalDays + 1);

                //Calculate the average CPD for the given period
                double AvgCPD = (Convert.ToDouble(dr[2]) * 1000) / TotalDays;

                //Add the CPD data to the list
                cpd[index] = new
                {
                    x = AvgCPD,
                    low = new DateTime(StartDay.Year, StartDay.Month, 1),
                    high = new DateTime(EndDay.Year, EndDay.Month, 1)
                };
                //cpd2[index] = new
                //{
                //    y = AvgCPD,
                //    low = new DateTime(StartDay.Year, StartDay.Month, 1),
                //    high = new DateTime(EndDay.Year, EndDay.Month, 1)
                //};
                index++;
            }

        }
        #endregion

        #region RAM Location Wise

        public void RAMLocWiseValues()
        {
            // string pathOfRAMLocWise = Directory.GetCurrentDirectory();
            DataTable d = s_RAMLocWise(Server.MapPath("RAM.xls"));
            int i = 0;
            List<Class1> klist = new List<Class1>();
            var sm1 = d.Select("location='(SM-1)'");
            var sm2 = d.Select("location='(SM-2)'");
            var sm3 = d.Select("location='(SM-3)'");
            var sm4 = d.Select("location='(SM-4)'");
            while (i < sm1.Length)
            {
                var f = sm1[i];
                Class1 f1 = new Class1();
                //sm1
                f1.start = (DateTime)f["StartDate"];
                f1.end = (DateTime)f["EndDate"];
                double days = (f1.start - f1.end).TotalDays;
                if (-(f1.start - f1.end).TotalDays < 181)
                {
                    f1.end = getEndDate_RAMLocWise(f1.start);
                }
                int row = i;
                Boolean err = false;
                int j = i;
                while ((DateTime)sm1[i]["EndDate"] < f1.end)
                {
                    row++;
                    i++;
                    if (i == sm1.Length)
                    {
                        err = true;
                        break;
                    }
                }
                if (err) break;
                f1.values = new Dictionary<string, double>();
                double ds1 = 0, ds2 = 0, ds3 = 0, ds4 = 0;
                i = j;


                for (; i <= row; i++)
                {
                    if (i == 0)
                    {
                        ds1 += getDose_RAMLocWise(sm1[i], f1.end, sm1[i]);
                        ds2 += getDose_RAMLocWise(sm2[i], f1.end, sm1[i]);
                        ds3 += getDose_RAMLocWise(sm3[i], f1.end, sm1[i]);
                        ds4 += getDose_RAMLocWise(sm4[i], f1.end, sm1[i]);
                    }
                    else
                    {
                        ds1 += getDose_RAMLocWise(sm1[i], f1.end, sm1[i - 1]);
                        ds2 += getDose_RAMLocWise(sm2[i], f1.end, sm1[i - 1]);
                        ds3 += getDose_RAMLocWise(sm3[i], f1.end, sm1[i - 1]);
                        ds4 += getDose_RAMLocWise(sm4[i], f1.end, sm1[i - 1]);
                    }
                }
                f1.values.Add("1", ds1);
                f1.values.Add("2", ds2);
                f1.values.Add("3", ds3);
                f1.values.Add("4", ds4);
                klist.Add(f1);
                i = j + 1;
                //SMvaluesDates.Add(kl)
            }

            //strEndDate = new string[klist.Count];
            //strStartDate = new string[klist.Count];
            loc1 = new dynamic[klist.Count];
            loc2 = new dynamic[klist.Count];
            loc3 = new dynamic[klist.Count];
            loc4 = new dynamic[klist.Count];
            for (int ii = 0; ii < klist.Count; ii++)
            {
                //Create dynamic lists here
                loc1[ii] = new { x = klist[ii].values["1"], low = klist[ii].start, high = klist[ii].end };
                loc2[ii] = new { x = klist[ii].values["2"], low = klist[ii].start, high = klist[ii].end };
                loc3[ii] = new { x = klist[ii].values["3"], low = klist[ii].start, high = klist[ii].end };
                loc4[ii] = new { x = klist[ii].values["4"], low = klist[ii].start, high = klist[ii].end };
                //Enddnamic lists
            }
        }
      



        public static double getDose_RAMLocWise(DataRow s1, DateTime end, DataRow s)
        {
            double abs, texpday;
            abs = (double)s1["AbsorbedDose"];
            texpday = (double)s1["TotalExpoDay"];
            double dose = (double)abs / (double)texpday;
            var sf = (DateTime)s1["startdate"];
            var sf1 = (DateTime)s["enddate"];
            double minusdays = 0;
            if (sf < sf1)
            {
                minusdays = (sf1 - sf).TotalDays;
            }
            var f1 = (DateTime)s1["EndDate"];
            if (end >= f1)
            {
                return Math.Round(abs, 2);
            }
            else
            {
                double d = -(((DateTime)s1["StartDate"] - end).TotalDays);
                return Math.Round(dose * (d - minusdays), 2);
            }
        }
        public static DateTime getEndDate_RAMLocWise(DateTime t)
        {
            DateTime d = t.AddDays(181);
            return d;
        }
        public static DataTable s_RAMLocWise(string path)
        {

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            String[] s = getsheet_RAMLocWise(path);
            for (int i = 0; i < s.Length; i++)
            {
                String query = "SELECT * FROM [" + s[i] + "]";
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
        private static String[] getsheet_RAMLocWise(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                String connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                  "Data Source=" + excelFile + ";Extended Properties=Excel 8.0;";

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
        #endregion

          #region class for RAM Values
        class Class1
        {
            public DateTime start;
            public DateTime end;
            public Dictionary<string, double> values;
        }
#endregion
    }
}