using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;


namespace WebApplication1.Controllers
{
    public class ExportToTimelineController : Controller
    {
        public void Index()
        {
			
			using (Presentation pres = new Presentation())
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Histogram, 50, 50, 500, 400);
				chart.ChartData.Categories.Clear();
				chart.ChartData.Series.Clear();

				IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

				wb.Clear(0);

				IChartSeries series = chart.ChartData.Series.Add(ChartType.Histogram);
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A1", 15));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A2", -41));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A3", 16));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A4", 10));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A5", -23));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A6", 16));

				chart.Axes.HorizontalAxis.AggregationType = AxisAggregationType.Automatic;

				pres.Save("C:/Users/esshreem/Histogram.pptx", SaveFormat.Pptx);
			}
		}

        public IActionResult Privacy()
        {
            return View();
        }

       
    }
    
}
