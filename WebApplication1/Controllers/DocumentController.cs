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
        private readonly ILogger<DocumentController> _logger;
        private readonly WordDocument wordDocument;

        public DocumentController(IWebHostEnvironment appEnvironment, ILogger<DocumentController> logger)
        {
            _appEnvironment = appEnvironment;
            _logger = logger;
            wordDocument = new WordDocument(HomeController.userFolder + HomeController.FileName);
        }

        //public async Task<IActionResult> SetWordDocument(string filename)
        //{
        //    wordDocument = new WordDocument(filename);
        //    return View("Index");
        //}

        public async Task<IActionResult> Index(string word, string text, string hyperlinkType, IFormFile image, int count = default)
        {

           


            wordDocument.CreateReferencesSection();
            string path = null;
            if (image != null)
            {
                path = _appEnvironment.WebRootPath + "/Images/" + image.FileName;
                using var fileStream = new FileStream(path, FileMode.Create);
                await image.CopyToAsync(fileStream);
            }

            if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(hyperlinkType))
            {
                if (hyperlinkType.Equals("bookmark"))
                {
                    wordDocument.CreateBookmarksForText(word, text, count);
                }
                else if (hyperlinkType.Equals("hyperlink"))
                {
                    wordDocument.CreateHyperlinksForText(word, text, count);
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
                    wordDocument.CreateBookmarksForImage(path, word, count);
                }
            }


            ViewBag.Messages = wordDocument.Messages;
            ViewBag.LongText = wordDocument.GetTextFromDocument();
            ViewBag.Footnotes = wordDocument.GetAllFootnotes();
            return View();
        }

        public IActionResult EditLinks()
        {
            var hyperlinks = wordDocument.GetAllHyperlinks().ToList();
            ViewBag.Hyperlinks = hyperlinks;
            var bookmarks = wordDocument.GetAllBookmarks().ToList();
            ViewBag.Bookmarks = bookmarks;
            var images = wordDocument.GetImages().ToList();
            ViewBag.Images = images;

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
            var list = wordDocument.GetAllHyperlinks().ToList();
            Field field = list[index];
            wordDocument.EditLinkInHypertext(field, hypertext);
            return Redirect("EditLinks");
        }

        public IActionResult DeleteHyperlink(int index)
        {
            var list = wordDocument.GetAllHyperlinks().ToList();
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

        public FileResult Download()
        {
            var file = new DirectoryInfo(HomeController.userFolder).GetFiles().Where(f => f.Name == HomeController.FileName).FirstOrDefault();
            var doc = new byte[0];
            doc = System.IO.File.ReadAllBytes(file.FullName);
            string fileExtension = file.Extension;
            string filename = file.Name;
            //DeleteFile(file);
            return File(doc, "application/" + fileExtension, filename);
        }

        private void DeleteFile(FileInfo file)
        {
            file.Delete();
        }


    }
}