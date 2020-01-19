using BusinessLogicLayer;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using WebApplication1.Models;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;

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
            //wordDocument.RemoveHyperlinks();
            StringBuilder longText = new StringBuilder();

            Document document = new Document();
            document.LoadFromFile("test.docx");

            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    longText.AppendLine($"{paragraph.Text}<br>");
                }
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
