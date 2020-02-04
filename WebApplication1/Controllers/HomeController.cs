using BusinessLogicLayer;
using Cyriller;
using Cyriller.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private const string path = "Files/Hypertext.docx";
        private readonly ILogger<HomeController> _logger;
        private Paragraph referencesParagraph;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index(string word, string text)
        {
            WordDocument wordDocument = new WordDocument("hyperlink");
            var paragraph = wordDocument.GetParagraphByWord("Сноски");
            if (paragraph == null)
            {
                Section sectionForReferences = wordDocument.Document.AddSection();
                var chapterForReferences = sectionForReferences.AddParagraph();
                chapterForReferences.AppendText("Сноски");
                chapterForReferences.AppendBreak(BreakType.LineBreak);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Heading1);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Title);

                //referencesParagraph = sectionForReferences.AddParagraph();
                referencesParagraph = chapterForReferences;
            }
            else
            {
                paragraph.AppendBreak(BreakType.LineBreak);
                referencesParagraph = paragraph;
            }

            
            

            if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word))
            {
                referencesParagraph.AppendText(text);
                referencesParagraph.AppendBreak(BreakType.LineBreak);
                wordDocument.SetReferencesWord(text.Split(' ')[0]);
                try
                {
                    CyrNounCollection cyrNounCollection = new CyrNounCollection();
                    CyrNoun cyrNoun = cyrNounCollection.Get(word, out CasesEnum @case, out NumbersEnum numbers);
                    var nounsSet = new HashSet<string>(cyrNoun.Decline().ToList());
                    var nouns = nounsSet.ToList();
                    if (cyrNoun.WordType != WordTypesEnum.Surname)
                    {
                        int nounLength = nouns.Count;
                        for (int i = 0; i < nounLength; i++)
                        {
                            nouns.Add(WordDocument.GetWordWithFirstLetterUpper(nouns[i]));
                        } 
                    }

                    foreach (var noun in nouns)
                    {
                        wordDocument.CreateBookmarks(noun, text);
                    }
                }
                catch (CyrWordNotFoundException error)
                {
                    ViewBag.Error = error.Message;
                }                
            }

            wordDocument.SaveCurrentDicument();

            ViewBag.LongText = wordDocument.GetTextFromDocument();
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
