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
       public static string excelFilepath = "";
       
        Computations c = new Computations();

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
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
        
        }

        
    
}