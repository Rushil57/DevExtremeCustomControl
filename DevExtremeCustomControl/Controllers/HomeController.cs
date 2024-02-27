using DevExtremeMvcApp1.Models;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Svg;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace DevExtremeMvcApp1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            


            return View();
        }

        [HttpPost]
        public void WriteImageToPPT(TestViewModel modelData)
        {
            string pictureFileName = "D:\\TestChartData_10.png";
            byte[] ImageBytes = Encoding.ASCII.GetBytes(modelData.ImageRawData.ToString().Split(',')[1]);
            //System.IO.File.WriteAllBytes("D:\\TestChartData_8.png", ImageBytes);
            System.IO.File.WriteAllBytes(pictureFileName, Convert.FromBase64String(modelData.ImageRawData.ToString().Split(',')[1]));

            

            Application pptApplication = new Application();

            Microsoft.Office.Interop.PowerPoint.Slides slides;
            Microsoft.Office.Interop.PowerPoint._Slide slide;
            Microsoft.Office.Interop.PowerPoint.TextRange objText;

            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            // Create new Slide
            slides = pptPresentation.Slides;
            slide = slides.AddSlide(1, customLayout);

            // Add title
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = "test";
            objText.Font.Name = "Arial";
            objText.Font.Size = 32;

            objText = slide.Shapes[2].TextFrame.TextRange;
            objText.Text = "";

            Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];
            slide.Shapes.AddPicture(pictureFileName, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, shape.Left, shape.Top, shape.Width, shape.Height);

            slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Test";

            pptPresentation.SaveAs(@"d:\test.pptx", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            //pptPresentation.Close();
            //pptApplication.Quit();
        }
    }
}