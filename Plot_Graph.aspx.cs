using DotNet.Highcharts.Helpers;
using DotNet.Highcharts.Options;
using System;
using System.Collections.Generic;
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
        string[] year_monthList;
        Object[] monthlyNumberList;
        Object[] smoothedNumberList;
        //List<string> year_monthList = new List<string>();
        //List<string> monthlyNUmberList = new List<string>();
        //List<string> smoothedNumberList = new List<string>();
        Computations c = new Computations();

        protected void Page_Load(object sender, EventArgs e)
        {
            loadSSNData();
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
            DotNet.Highcharts.Highcharts chart = new DotNet.Highcharts.Highcharts("chart").InitChart(new Chart { ZoomType = DotNet.Highcharts.Enums.ZoomTypes.X })
               .SetXAxis(new[]{
                new XAxis
                            {
                               Id="SunSpot_Axis",
                                Categories = year_monthList,
                            }
                           
                })
               .SetSeries(new[]
                { new Series
                            {
                                XAxis="SunSpot_Axis",
                                Name="First Series",
                                Data = new Data(monthlyNumberList)
                            },
                    new Series
                            {
                                XAxis="SunSpot_Axis",
                                Name="Second series",
                                Data = new Data(smoothedNumberList)
                            }
                });


            ltrChart.Text = chart.ToHtmlString();

        }


        public void loadSSNData()
        {
            try
            {

                string str = "";
                // string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Data\Names.txt");
                string pathOfSSNFile = Server.MapPath("Sunspot_Dataset.xlsx");

                Excel.Range range;
                MyApp = new Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(pathOfSSNFile); //Giving the path to excel workbook to open and read the file.
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];
                int rCnt, cCnt, YearMonth = 0, MNUmb = 0, SmNumb = 0;
                range = MySheet.UsedRange;

                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    str = Convert.ToString((range.Cells[1, cCnt] as Excel.Range).Value2);
                    if (!String.IsNullOrEmpty(str))
                    {
                        if (str.Contains("Year_Month")) // Storing the column number in dateCol which has Date in the first row
                        {
                            YearMonth = cCnt;

                        }
                        if (str.Contains("Monthly Number")) // Storing the column number in doseCol which has Total in the first row.
                        {

                            MNUmb = cCnt;
                        }
                        if (str.Contains("SmoothedNumber"))
                        {
                            SmNumb = cCnt;

                        }
                    }
                }
                year_monthList = new string[range.Rows.Count];
                monthlyNumberList = new Object[range.Rows.Count];
                smoothedNumberList = new Object[range.Rows.Count];
                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    string yearMonth_Data = (range.Cells[rCnt, YearMonth] as Excel.Range).Value2;
                    object monthlyNumberData = (range.Cells[rCnt, MNUmb] as Excel.Range).Value2;
                    object smootherNumberData = (range.Cells[rCnt, SmNumb] as Excel.Range).Value2;

                    year_monthList[rCnt - 1] = yearMonth_Data;
                    monthlyNumberList[rCnt - 1] = monthlyNumberData;
                    smoothedNumberList[rCnt - 1] = smootherNumberData;

                }

            }
            catch (Exception e) { return; }
        }
       

    }



}