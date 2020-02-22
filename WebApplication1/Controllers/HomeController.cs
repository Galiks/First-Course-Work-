using BusinessLogicLayer;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc.Fields;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        readonly IWebHostEnvironment _appEnvironment;
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
            string path = null;
            if (image != null)
            {
                path = _appEnvironment.WebRootPath + "/Files/" + image.FileName;
                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    await image.CopyToAsync(fileStream);
                }
            }

            if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("bookmark"))
                {
                    wordDocument.CreateBookmarksForText(word, text);
                }
                if (hyperlinkType.Equals("hyperlink"))
                {
                    wordDocument.CreateHyperlinksForText(word, text);
                }
            }
            else if (!string.IsNullOrWhiteSpace(path) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("bookmark"))
                {
                    wordDocument.CreateBookmarksForImage(path, word);
                }
                if (hyperlinkType.Equals("hyperlink"))
                {
                    wordDocument.CreatHyperlinkForImage(path, text);
                }
            }


            ViewBag.Messages = wordDocument.Messages;
            ViewBag.LongText = wordDocument.GetTextFromDocument();
            ViewBag.Footnotes = wordDocument.GetAllFootnotes();
            return View();
        }

        public IActionResult EditLinks()
        {
            ViewBag.Hyperlinks = wordDocument.GetAllHyperlinks();
            var bookmarks = wordDocument.GetAllBookmarks();
            ViewBag.Bookmarks = bookmarks.ToList();

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

        public IActionResult UpdateHyperlink(int index, string hypertext)
        {
            List<Field> list = wordDocument.GetAllHyperlinks();
            Field field = list[index];
            wordDocument.EditLinkInHypertext(field, hypertext);
            return Redirect("EditLinks");
        }

        public IActionResult DeleteHyperlink(int index)
        {
            List<Field> list = wordDocument.GetAllHyperlinks();
            Field field = list[index];
            this.wordDocument.DeleteHyperlink(field);
            return Redirect("EditLinks");
        }

        public IActionResult UpdateBookmark(string bookmark, string text)
        {
            wordDocument.EditTextInBookmark(bookmark, text);
            return Redirect("EditLinks");
        }

        public IActionResult DeleteBookmark(string bookmark)
        {
            wordDocument.DeleteBookmark(bookmark);
            return Redirect("EditLinks");
        }
    }
}
