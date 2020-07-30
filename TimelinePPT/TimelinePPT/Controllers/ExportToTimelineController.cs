using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;
using Spire.Presentation;
using Spire.Presentation.Charts;
using System.Collections;
using System.IO;
using System.Net.Mime;
using Microsoft.AspNetCore.Cors;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;

namespace TimelinePPT.Controllers
{



    public class SpireController : Controller
    {

        [EnableCors("MyPolicy")]


       /** [HttpGet]
        [Route("api/FileAPI/GetFile")]
        public HttpResponseMessage GetFile(string fileName)
        {
            //Create HTTP Response.
            HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);
            //Set the File Path.
            string filePath = HttpContext.Current.Server.MapPath("~/Files/") + fileName;
            //Check whether File exists.
            if (!File.Exists(filePath))
            {
                //Throw 404 (Not Found) exception if File not found.
                response.StatusCode = HttpStatusCode.NotFound;
                response.ReasonPhrase = string.Format("File not found: {0} .", fileName);
                throw new HttpResponseException(response);
            }
            //Read the File into a Byte Array.
            byte[] bytes = File.ReadAllBytes(filePath);
            //Set the Response Content.
            response.Content = new ByteArrayContent(bytes);
            //Set the Response Content Length.
            response.Content.Headers.ContentLength = bytes.LongLength;
            //Set the Content Disposition Header Value and FileName.
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = fileName;
            //Set the File Content Type.
            response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(fileName));
            return response;
        }**/

        public HttpResponseMessage Index([System.Web.Http.FromBody] Task2[] taskData)
        {




            using (Presentation presentation = new Presentation())
            {

                Task t0 = new Task("task0", new DateTime(2020, 01, 01), new DateTime(2020, 01, 15));
                Task t1 = new Task("task1", new DateTime(2020, 02, 05), new DateTime(2020, 02, 29));
                Task t2 = new Task("task2", new DateTime(2020, 03, 10), new DateTime(2020, 03, 30));
                Task t3 = new Task("task3", new DateTime(2020, 04, 10), new DateTime(2020, 04, 30));
                Task t4 = new Task("task4", new DateTime(2020, 05, 05), new DateTime(2020, 05, 26));
                Task t5 = new Task("task5", new DateTime(2020, 06, 10), new DateTime(2020, 06, 30));
                Task t6 = new Task("task6", new DateTime(2020, 07, 10), new DateTime(2020, 07, 28));
                Task t7 = new Task("task7", new DateTime(2020, 07, 10), new DateTime(2020, 07, 30));
                Task t8 = new Task("task8", new DateTime(2020, 08, 01), new DateTime(2020, 08, 01));
                Task t9 = new Task("task9", new DateTime(2020, 08, 05), new DateTime(2020, 08, 22));
                Task t10 = new Task("task10", new DateTime(2020, 09, 10), new DateTime(2020, 09, 30));

                var timelineTableDimensions = new Double[30];




                var horizontalAxes = new Double[30];
                var slideWidth = presentation.SlideSize.Size.Width;
                var rectangleOffset = 30;

                var oneWidth = (slideWidth - rectangleOffset * 2) / 12;
                string[] monthTitles = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec" };
                int[] days = { 31, 29, 31, 30, 31, 30, 30, 31, 30, 31, 30, 31 };

                var slideHeight = presentation.SlideSize.Size.Height;
                //  string htmlText = "<html><body>";
                var XAxis = 30;
                float chartWidth = (int)slideWidth - rectangleOffset * 2;
                ArrayList list = new ArrayList();
                IAutoShape RectangleContainer1 = presentation.Slides[0].Shapes.AppendShape(ShapeType.TwoDiagonalRoundCornerRectangle, new RectangleF(XAxis, 70, chartWidth, 30));
                // XAxis = XAxis + (int)oneWidth;
                //The first rectangle should have a rounded corner
                RectangleContainer1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                RectangleContainer1.Fill.SolidColor.Color = Color.DarkOrange;
                RectangleContainer1.ShapeStyle.LineColor.Color = Color.DarkOrange;
                TextRange textRange = RectangleContainer1.TextFrame.TextRange;

                float singleDayWidth = chartWidth / 366;
                for (int j = 0; j < 12; j++)
                {
                    IAutoShape rectangleContainer2 = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new Rectangle(XAxis, 75, (int)oneWidth, 20));
                    XAxis = XAxis + (int)oneWidth;
                    rectangleContainer2.ShapeStyle.FillColor.Color = Color.Transparent;
                    textRange = rectangleContainer2.TextFrame.TextRange;
                    rectangleContainer2.AppendTextFrame(monthTitles[j]);
                    textRange.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    rectangleContainer2.ShapeStyle.LineColor.Color = Color.White;
                    textRange.Fill.SolidColor.Color = Color.Green;
                    rectangleContainer2.Fill.SolidColor.Color = Color.Yellow;
                    rectangleContainer2.TextFrame.TextRange.FontHeight = 11;

                    //set the Font of text in shape
                    textRange.FontHeight = 12;
                    textRange.IsItalic = TriState.True;
                    textRange.TextUnderlineType = TextUnderlineType.Single;
                    textRange.LatinFont = new TextFont("Gulim");
                    list.Add(rectangleContainer2);

                }
                // The last rectangle should have a rounded corner

                presentation.Slides[0].GroupShapes(list);

                //   RectangleContainer2.TextFrame.Paragraphs.AddFromHtml(htmlText);

                /* IAutoShape RectangleContainer = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new Rectangle(rectangleOffset, 0, (int)(slideWidth - rectangleOffset * 2), (int)slideHeight - rectangleOffset * 2));
                 // RectangleContainer.Fill.SolidColor.Color = Color.Yellow;
                 RectangleContainer1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                 RectangleContainer1.Fill.SolidColor.Color = Color.DarkOrange;
                 RectangleContainer1.ShapeStyle.LineColor.Color = Color.DarkOrange;
                 RectangleContainer.ShapeStyle.FillColor.Color = Color.Transparent;

                 RectangleContainer.Line.SolidFillColor.Color = Color.Black;*/

                ArrayList tasks = new ArrayList();
                tasks.Add(t0);
                tasks.Add(t1);
                tasks.Add(t2);
                tasks.Add(t3);
                tasks.Add(t4);
                tasks.Add(t5);
                tasks.Add(t6);
                tasks.Add(t7);
                tasks.Add(t8);
                tasks.Add(t9);
                tasks.Add(t10);
                int initialOffsetX = 30;
                int initialOffsetY = 100;
                int initialOffsetW = 0;
                int initialOffsetH = 15;
                //int widthOfSlide = presentation.Slides[0].
                int i = 0;
                /* IAutoShape tl = presentation.Slides[0].Shapes.AppendShape(ShapeType.TwoDiagonalRoundCornerRectangle, new RectangleF(initialOffsetX, 50, (float)(365*1.5), 20));
                 tl.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                 tl.Fill.SolidColor.Color = .DarkOrange;
                 tl.ShapeStyle.LineColor.Color = Color.DarkOrange;
                 tl.ShapeStyle.FontColor.Color = Color.White;
                 tl.TextFrame.TextRange.FontHeight = 11;
                 tl.AppendTextFrame("JAN | FEB | MAR | APR | MAY | JUN | JUL | AUG | SEP | OCT | NOV | DEC");*/
                foreach (Task task in tasks)
                {
                    i++;
                    int stMonth = task.startDate.Month;
                    int enMonth = task.endDate.Month;
                    TimeSpan ts = task.endDate - task.startDate;
                    int stDay = task.startDate.Day;
                    int enDay = task.endDate.Day;
                    float startDay = (stMonth - 1) * (days[stMonth - 1] * singleDayWidth) + stDay;
                    float x = initialOffsetX + startDay;
                    float y = initialOffsetY + i * 20;
                    double totalDays = ts.TotalDays;
                    if (totalDays == 0)
                    {
                        var shape1 = presentation.Slides[0].Shapes.AppendShape(ShapeType.RightTriangle, new RectangleF(x, 35, 15, 20));
                        var shape2 = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(x, 35, 1, 40));
                        shape1.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                        shape1.Fill.SolidColor.Color = Color.OrangeRed;
                        shape1.ShapeStyle.LineColor.Color = Color.OrangeRed;
                        shape2.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                        shape2.Fill.SolidColor.Color = Color.OrangeRed;
                        shape2.ShapeStyle.LineColor.Color = Color.OrangeRed;
                        ArrayList list2 = new ArrayList() { shape1, shape2 };
                        presentation.Slides[0].GroupShapes(list2);
                        continue;
                    }

                    float w = (float)(initialOffsetW + (totalDays * singleDayWidth));
                    int h = initialOffsetH;

                    IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(x, y, w, h));
                    shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                    shape.Fill.SolidColor.Color = Color.DarkOrange;
                    shape.ShapeStyle.LineColor.Color = Color.DarkOrange;
                    shape.ShapeStyle.FontColor.Color = Color.Black;
                    shape.TextFrame.TextRange.FontHeight = 5;
                    shape.AppendTextFrame(stMonth + "/" + stDay + "-" + enMonth + "/" + enDay);
                    IAutoShape rc = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF((float)(x + w + 2), y, 40, h));
                    rc.ShapeStyle.FillColor.Color = Color.Transparent;
                    rc.ShapeStyle.LineColor.Color = Color.Transparent;
                    rc.ShapeStyle.FontColor.Color = Color.OrangeRed;
                    rc.TextFrame.TextRange.FontHeight = 10;
                    rc.AppendTextFrame(task.name);

                    //string htmlText = "<html><body><div style='font-size:11.0pt'>"+task.name+"</div></body></html>";
                    //rc.TextFrame.Paragraphs.AddFromHtml(htmlText);
                }

                /*  IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 0, 50, 10));
                  shape.Name = "Task1";

                  //shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                  //shape.Fill.SolidColor.Color = Color.LightGreen;
                  //shape.ShapeStyle.LineColor.Color = Color.White;
                  shape.TextFrame.Text = "Task33";


                  IAutoShape shape2 = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 20, 50, 10));
                  shape2.Name = "Task2";

                  IAutoShape shape3 = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 40, 50, 10));
                  shape3.Name = "Task3";*//*  IAutoShape shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 0, 50, 10));
                  shape.Name = "Task1";

                  //shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
                  //shape.Fill.SolidColor.Color = Color.LightGreen;
                  //shape.ShapeStyle.LineColor.Color = Color.White;
                  shape.TextFrame.Text = "Task33";


                  IAutoShape shape2 = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 20, 50, 10));
                  shape2.Name = "Task2";

                  IAutoShape shape3 = presentation.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, new RectangleF(50, 40, 50, 10));
                  shape3.Name = "Task3";*/

                // Write the presentation file to disk
                //System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();



                //System.IO.Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                /**  byte[] buffer = new Byte[1000];

                  using (MemoryStream fs = new FileStream())
                  {
                      // ...
                      presentation.SaveToFile(fs, FileFormat.Pptx2010);

                      var result = HttpContext.Response;
                     
                       result.ContentType= "application/vnd.ms-powerpoint";

                      result.Headers.Add("Content-Disposition", "attachment");
                      int length;
                      do
                      {
                          // Verify that the client is connected.
                          if (!HttpContext.RequestAborted.IsCancellationRequested)
                          {
                              // Read data into the buffer.
                              length = fs.Read(buffer, 0, 1000);

                              // and write it out to the response's output stream
                              result.Body.WriteAsync(buffer, 0, length);


                              //Clear the buffer
                              buffer = new Byte[1000];
                          }
                          else
                          {
                              // cancel the download if client has disconnected
                              length = -1;
                          }
                      } while (length > 0);


                      return result;



                  }**/
                var filePath = "To_stream.pptx";
                presentation.SaveToFile(filePath, FileFormat.Pptx2010);

                //  byte[] buffer = new Byte[1000];

                //System.IO.File fs = new File("To_stream.pptx");
                //  byte[] buffer = new Byte[1000];
                byte[] buffer = System.IO.File.ReadAllBytes(filePath);




                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                result.Content = new ByteArrayContent(buffer);


                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-powerpoint");


                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                result.Content.Headers.ContentLength = buffer.Length;
           

                return result;




            }



        }


        
        public class Task
        {
            public string name { get; set; }
            public DateTime startDate { get; set; }
            public DateTime endDate { get; set; }
            public Task(string name, DateTime stDate, DateTime enDate)
            {
                this.name = name;
                this.startDate = stDate;
                this.endDate = enDate;
            }
        }
    }
}