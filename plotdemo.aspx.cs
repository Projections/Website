using DotNet.Highcharts.Helpers;
using DotNet.Highcharts.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Projections_Capstone_Spring15
{
    public partial class plotdemo : System.Web.UI.Page
    {



        protected void Page_Load(object sender, EventArgs e)
        {
            var x = new[]{ new {x = 1.0, low = new DateTime(2005,03,15), high = new DateTime(2005,07,20) }, 
                           new {x = 1.5, low = new DateTime(2006,01,09), high = new DateTime(2006,03,12)}
            };

            dynamic y = new dynamic[4];
            for(int i=0;i<3;i++)
            {
                y[i]=new { x=i,low = new DateTime(2005,03,15), high = new DateTime(2005,07,20)};
            }

             DotNet.Highcharts.Highcharts RAMChart = new DotNet.Highcharts.Highcharts("chart1").InitChart(new Chart
            {
                ZoomType = DotNet.Highcharts.Enums.ZoomTypes.X,
                Type = DotNet.Highcharts.Enums.ChartTypes.Columnrange,  
                Inverted=true   
            })
            .SetXAxis(new[]{
                new XAxis
                            {
                             //Id="RAM Axes",
                               Type=DotNet.Highcharts.Enums.AxisTypes.Linear,
                               // Categories = new[]{"Jan","Feb","Mar"},
                                //Labels=new XAxisLabels{Step=10, StaggerLines=1}
                               
                             // MinRange=30*24
                            }

        })
        .SetYAxis(new[]{
        new YAxis{
             Type=DotNet.Highcharts.Enums.AxisTypes.Datetime
            }
        })
        ;
             RAMChart.SetSeries(new[]
                {                             
                             new Series
                            {
                                Name="SM-4",
                                Data=new Data(y)
                            },
                            new Series
                            {
                                Name="SM-4",
                                Data=new Data(x)
                            }
                });
            ltrPlot.Text = RAMChart.ToHtmlString();
        }
        
    }
}