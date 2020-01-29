using BusinessLogicLayer;
using Cyriller;
using Cyriller.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Diagnostics;
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
            WordDocument wordDocument = new WordDocument("test");

            StringBuilder longText = new StringBuilder();

            Document document = new Document();
            //document.LoadFromFile("hyperlink.docx.docx");
            document.LoadFromFile("test.docx");

            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    StringBuilder paragraphText = new StringBuilder(paragraph.Text);

                    foreach (DocumentObject child in paragraph.ChildObjects)
                    {
                        if (child.DocumentObjectType == DocumentObjectType.Field)
                        {
                            Field field = child as Field;
                            if (field.Type == FieldType.FieldHyperlink & !string.IsNullOrWhiteSpace(field.FieldText))
                            {
                                paragraphText.Replace(field.FieldText, $"<strong>{field.FieldText}</strong>");
                            }
                        }
                    }

                    longText.AppendLine($"{paragraphText}<br>");
                }
            }

            ViewBag.LongText = longText;
            if (!string.IsNullOrWhiteSpace(text))
            {
                try
                {
                    
                    CyrNounCollection cyrNounCollection = new CyrNounCollection();
                    CyrNoun noun = cyrNounCollection.Get(text, out CasesEnum @case, out NumbersEnum numbers);
                }
                catch (CyrWordNotFoundException error)
                {
                    ViewBag.Error = error.Message;
                }
                
            }
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
