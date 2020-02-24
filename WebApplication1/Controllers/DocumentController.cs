﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using BusinessLogicLayer;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc.Fields;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class DocumentController : Controller
    {
        readonly IWebHostEnvironment _appEnvironment;
        private readonly ILogger<HomeController> _logger;
        private WordDocument wordDocument;
        private readonly List<string> formats;

        public DocumentController(IWebHostEnvironment appEnvironment, ILogger<HomeController> logger)
        {
            _appEnvironment = appEnvironment;
            _logger = logger;
            formats = new List<string>() { ".docx", ".pdf", ".doc" };
            wordDocument = new WordDocument(HomeController.filename);
        }

        public async Task<IActionResult> SetWordDocument(string filename)
        {
            wordDocument = new WordDocument(filename);
            return View("Index");
        }

        public async Task<IActionResult> Index(string word, string text, string hyperlinkType, IFormFile image, IFormFile file)
        {
            
            wordDocument.CreateReferencesSection();
            string path = null;
            if (image != null)
            {
                path = _appEnvironment.WebRootPath + "/Files/Images/" + image.FileName;
                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    await image.CopyToAsync(fileStream);
                }
                var extension = Path.GetExtension(path);

            }

            if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("bookmark"))
                {
                    wordDocument.CreateBookmarksForText(word, text);
                }
                else if (hyperlinkType.Equals("hyperlink"))
                {
                    wordDocument.CreateHyperlinksForText(word, text);
                }
            }
            else if (!string.IsNullOrWhiteSpace(path) & !string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("hyperlink"))
                {
                    wordDocument.CreatHyperlinkForImage(path, text);
                }
            }
            else if (!string.IsNullOrWhiteSpace(path) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("bookmark"))
                {
                    wordDocument.CreateBookmarksForImage(path, word);
                }
            }


            ViewBag.Messages = wordDocument.Messages;
            ViewBag.LongText = wordDocument.GetTextFromDocument();
            ViewBag.Footnotes = wordDocument.GetAllFootnotes();
            return View();
        }

        public IActionResult EditLinks()
        {
            List<Field> hyperlinks = wordDocument.GetAllHyperlinks();
            ViewBag.Hyperlinks = hyperlinks;
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

        public IActionResult UpdateHyperlink(int index, string hypertext, IFormFile image)
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

        public IActionResult UpdateBookmark(string bookmark, string text, IFormFile image)
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