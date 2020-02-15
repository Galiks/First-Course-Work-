using BusinessLogicLayer;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        IWebHostEnvironment _appEnvironment;
        private readonly ILogger<HomeController> _logger;
        //private Paragraph referencesParagraph;
        //private Section referencesSection;
        private readonly WordDocument wordDocument;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment appEnvironment)
        {
            _logger = logger;
            wordDocument = new WordDocument("hyperlink");
            _appEnvironment = appEnvironment;
        }

        public async Task<IActionResult> Index(string word, string text, string hyperlinkType, IFormFile image)
        {

            wordDocument.CreateReferencesSection();

            if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("bookmark"))
                {
                    wordDocument.CreateBookmarks(word, text);
                }
                if (hyperlinkType.Equals("hyperlink"))
                {
                    wordDocument.CreateHyperlinksForText(word, text);
                }
            }

            if (image != null)
            {
                string path = "/Files/" + image.FileName;
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await image.CopyToAsync(fileStream);
                }
            }

            ViewBag.Messages = wordDocument.Messages;
            ViewBag.LongText = wordDocument.GetTextFromDocument();
            ViewBag.Footnotes = wordDocument.GetAllFootnotes();
            return View();
        }



        public IActionResult EditLinks()
        {
            Section section = wordDocument.GetSectionAndParagraphByWord("Сноски").Item1;
            //ViewBag.Links = wordDocument.FindAllBookmarkBySection(section);
            ViewBag.Hyperlinks = wordDocument.FindAllLinks();
            ViewBag.Bookmarks = wordDocument.FindAllBookmarkBySection(section);

            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult UsingHypertext()
        {
            return View();
        }
    }
}
