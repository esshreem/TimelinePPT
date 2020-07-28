using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using System.Drawing;

namespace TimelinePPT.Controllers
{
    
    public class ValuesController : Controller
    {
        public void Index()
        {
         

            using (Presentation presentation = new Presentation())
            {
                // Adding chart
                IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.StackedBar, 50, 50, 600, 400, true);
                IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

                IChartCategoryCollection categories = chart.ChartData.Categories;
                categories[0].Value = "cat1";
                int defaultWorksheetIndex = 0;

                chart.ChartData.Series.Clear();
                chart.ChartData.Categories.Clear();
                int s = chart.ChartData.Series.Count;
                s = chart.ChartData.Categories.Count;

                // Adding new series
                chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.Type);
                chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.Type);

                // Adding new categories
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 6, 0, "Task 1"));
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 5, 0, "Task 2"));
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 4, 0, "Task 3"));
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 3, 0, "Task 1"));
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 2, 0, "Task 2"));
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 1, 0, "Task 3"));


                // Take first chart series
                IChartSeries series = chart.ChartData.Series[0];

                // Now populating series data

                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 3));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 6));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 11));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 4, 1, 3));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 5, 1, 6));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 6, 1, 11));
                // Setting fill color for series
                series.Format.Fill.FillType = FillType.Solid;
                series.Format.Fill.SolidFillColor.Color = Color.Transparent;


                // Take second chart series
                series = chart.ChartData.Series[1];

                // Now populating series data
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 2, 1));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 2, 4));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 2, 1));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 4, 2, 1));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 5, 2, 4));
                series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 6, 2, 1));

                // Setting fill color for series
                series.Format.Fill.FillType = FillType.Solid;
                series.Format.Fill.SolidFillColor.Color = System.Drawing.Color.Green;
            

                // First label will be show Category name
                IDataLabel lbl = series.DataPoints[0].Label;
                lbl.DataLabelFormat.ShowCategoryName = true;

                lbl = series.DataPoints[1].Label;
                lbl.DataLabelFormat.ShowSeriesName = true;

                // Show value for third label
                lbl = series.DataPoints[2].Label;
                lbl.DataLabelFormat.ShowValue = true;
                lbl.DataLabelFormat.ShowSeriesName = true;
                lbl.DataLabelFormat.Separator = "/";

                // Write the presentation file to disk
                presentation.Save("C:/Users/esshreem/Histogram.pptx", SaveFormat.Pptx);
            }
        }
    }
}
