using BusinessLogicLayer;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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

            //if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            //{
            //    if (hyperlinkType.Equals("bookmark"))
            //    {
            //        wordDocument.CreateBookmarks(word, text);
            //    }
            //    if (hyperlinkType.Equals("hyperlink"))
            //    {
            //        wordDocument.CreateHyperlinksForText(word, text);
            //    }
            //}

            if (image != null)
            {
                string path = _appEnvironment.WebRootPath + "/Files/" + image.FileName;
                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    await image.CopyToAsync(fileStream);
                }

                //wordDocument.CreatHyperlinkForImage(path, text);

                Document document = new Document();
                document.LoadFromFile("Test.docx");

                Section section = document.Sections[0];
                section.Paragraphs[0].AppendText("ПРИВЕТ! КАК ДЕЛА?!");
                section.Paragraphs[0].AppendBookmarkStart("ПРИВЕТ!");
                section.Paragraphs[0].AppendBookmarkEnd("КАК ДЕЛА?!");


                BookmarksNavigator bn = new BookmarksNavigator(document);
                bn.MoveToBookmark("ПРИВЕТ!", true, true);
                Section section0 = document.AddSection();
                Paragraph paragraph = section0.AddParagraph();
                Image image2 = Image.FromFile(path);
                DocPicture picture = paragraph.AppendPicture(image2);
                bn.InsertParagraph(paragraph);
                document.Sections.Remove(section0);
                string output = "sample.docx";
                document.SaveToFile(output, FileFormat.Docx);

                //wordDocument.CreateBookmarksForImage(path, word);
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
