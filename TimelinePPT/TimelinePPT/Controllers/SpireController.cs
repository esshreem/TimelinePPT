using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;

namespace TimelinePPT.Controllers
{

    public class SpireController : Controller
    {
        public void Index()
        {

            var timelinetype = "Year";
            using (Presentation presentation = new Presentation())
            {
                var timelineTableDimensions = new Double[30];
                // Adding chart
                if( timelinetype == "Year")
                {
                    timelineTableDimensions = new Double[12];

                }

                IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.TwoDiagonalRoundCornerRectangle, new RectangleF(50, 0, 100, 100));
                
                ITable table2 = presentation.Slides[0].Shapes.AppendTable(20, 30, new Double[2] { 5, 5 }, new Double[] { 10 });
                shape.Name = "Task1";
            
                IAutoShape shape2 = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 150, 100, 100));
                shape2.Name = "Task2";

                IAutoShape shape3 = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 300, 100, 100));
                shape3.Name = "Task3";

                // IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.StackedBar, 50, 50, 600, 400, true);



                // Write the presentation file to disk
                presentation.SaveToFile("C:/Users/esshreem/Histogram.pptx", FileFormat.Pptx2010);

            }
        }
    }
}
