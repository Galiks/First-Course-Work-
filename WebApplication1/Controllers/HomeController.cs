using BusinessLogicLayer;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private const string path = "Files/Hypertext.docx";
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index(string text)
        {
            WordDocument wordDocument = new WordDocument();
            //wordDocument.CreateDocument();
            //wordDocument.Result();
            StringBuilder longText = new StringBuilder();

            WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open("test.docx", false);
            IEnumerable<Paragraph> paragraphElements =
                wordProcessingDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>();

            foreach (Paragraph p in paragraphElements)
            {

                //IEnumerable<Text> textElements = p.Descendants<Text>();

                //foreach (Text t in textElements)
                //{
                //   longText.Append(t.Text);
                //}

                longText.AppendLine(p.InnerText);
                //longText.AppendLine();
            }

            ViewBag.LongText = longText;
            ViewBag.Text = text;
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult TextRedactor()
        {
            return View();
        }
    }
}
