using BusinessLogicLayer;
using Cyriller;
using Cyriller.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private const string path = "Files/Hypertext.docx";
        private readonly ILogger<HomeController> _logger;
        private Paragraph referencesParagraph;
        private Section referencesSection;
        private readonly WordDocument wordDocument;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
            wordDocument = new WordDocument("hyperlink");
        }

        public IActionResult Index(string word, string text, string hyperlinkType)
        {
            var sectionAndParagraph = wordDocument.GetSectionAndParagraphByWord("Сноски");
            if (sectionAndParagraph == null)
            {
                Section sectionForReferences = wordDocument.Document.AddSection();
                var chapterForReferences = sectionForReferences.AddParagraph();
                chapterForReferences.AppendText("Сноски");
                chapterForReferences.AppendBreak(BreakType.LineBreak);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Heading1);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Title);

                //referencesParagraph = sectionForReferences.AddParagraph();
                referencesParagraph = chapterForReferences;
                referencesSection = sectionForReferences;
            }
            else
            {
                referencesParagraph = sectionAndParagraph.Item2;
                referencesSection = sectionAndParagraph.Item1;
            }

            wordDocument.SaveCurrentDicument();

            if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("bookmark"))
                {
                    wordDocument.SetReferencesWord(text);
                    Paragraph newParagraph = referencesSection.AddParagraph();

                    try
                    {
                        CyrNounCollection cyrNounCollection = new CyrNounCollection();
                        CyrNoun cyrNoun = cyrNounCollection.Get(word, out CasesEnum @case, out NumbersEnum numbers);
                        var nounsSet = new HashSet<string>(cyrNoun.Decline().ToList());
                        foreach (var noun in cyrNoun.DeclinePlural().ToList())
                        {
                            nounsSet.Add(noun);
                        }
                        var nouns = nounsSet.ToList();
                        if (cyrNoun.WordType != WordTypesEnum.Surname)
                        {
                            int nounLength = nouns.Count;
                            for (int i = 0; i < nounLength; i++)
                            {
                                nouns.Add(WordDocument.GetWordWithFirstLetterUpper(nouns[i]));
                            }
                        }

                        for (int i = 0; i < nouns.Count; i++)
                        {
                            string noun = nouns[i];
                            if (i == 0)
                            {
                                wordDocument.CreateBookmarks(noun, newParagraph, text);
                            }
                            else
                            {
                                wordDocument.CreateBookmarks(noun, newParagraph);
                            }
                        }
                    }
                    catch (CyrWordNotFoundException error)
                    {
                        ViewBag.Error = error.Message;
                        wordDocument.CreateBookmarks(word, newParagraph, text);
                    }
                    finally
                    {
                        if (wordDocument.IndexNextField.Equals(default))
                        {
                            wordDocument.IncreaseOfTwoindexNextField();
                        }
                    }
                }

                if (hyperlinkType.Equals("hyperlink"))
                {
                    try
                    {
                        CyrNounCollection cyrNounCollection = new CyrNounCollection();
                        CyrNoun cyrNoun = cyrNounCollection.Get(word, out CasesEnum @case, out NumbersEnum numbers);
                        var nounsSet = new HashSet<string>(cyrNoun.Decline().ToList());
                        foreach (var noun in cyrNoun.DeclinePlural().ToList())
                        {
                            nounsSet.Add(noun);
                        }
                        var nouns = nounsSet.ToList();
                        if (cyrNoun.WordType != WordTypesEnum.Surname)
                        {
                            int nounLength = nouns.Count;
                            for (int i = 0; i < nounLength; i++)
                            {
                                nouns.Add(WordDocument.GetWordWithFirstLetterUpper(nouns[i]));
                            }
                        }

                        for (int i = 0; i < nouns.Count; i++)
                        {
                            string noun = nouns[i];
                            wordDocument.CreateHyperlinkByWord(noun, text);
                        }
                    }
                    catch (CyrWordNotFoundException error)
                    {
                        ViewBag.Error = error.Message;
                        wordDocument.CreateHyperlinkByWord(word, text);
                    }
                    finally
                    {
                        if (wordDocument.IndexNextField.Equals(default))
                        {
                            wordDocument.IncreaseOfTwoindexNextField();
                        }
                    }
                }
            }

            ViewBag.Messages = wordDocument.Messages;
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

        public IActionResult UsingHypertext()
        {
            return View();
        }
    }
}
