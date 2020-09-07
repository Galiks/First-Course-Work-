using BusinessLogicLayer;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc.Fields;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
            _logger.LogInformation($"В класс WordDocument передан файл {HomeController.filepath}.");
            wordDocument = new WordDocument(HomeController.filepath);
        }

        public async Task<IActionResult> Index(string word, string text, string linkType, IFormFile image, int count = default)
        {
            string path = null;
            if (image != null)
            {
                path = _appEnvironment.WebRootPath + "/Images/" + image.FileName;
                try
                {
                    await FolderWork.SaveImage(image, path);
                    _logger.LogInformation($"Изображение {image.FileName} сохранено по пути {path}.");
                }
                catch (System.Exception)
                {
                    _logger.LogError($"Не удалось сохранить изображение {image.FileName}.");
                }
            }

            if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(linkType))
            {
                if (linkType.Equals("bookmark"))
                {
                    _logger.LogInformation($"Начато создание закладки для текста с параметрами {text}, {word}");
                    wordDocument.CreateBookmarksForText(word, text, count);
                }
                else if (linkType.Equals("hyperlink"))
                {
                    _logger.LogInformation($"Начато создание гиперссылки для текста с параметрами {text}, {word}");
                    wordDocument.CreateHyperlinksForText(word, text, count);
                }
            }
            else if (!string.IsNullOrWhiteSpace(path) & !string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(linkType))
            {
                if (linkType.Equals("hyperlink"))
                {
                    _logger.LogInformation($"Начато создание гиперссылки для изображения с параметрами {text}, {path}");
                    wordDocument.CreatHyperlinkForImage(path, text);
                }
            }
            else if (!string.IsNullOrWhiteSpace(path) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(linkType))
            {
                if (linkType.Equals("bookmark"))
                {
                    _logger.LogInformation($"Начато создание закладки для изображения с параметрами {text}, {path}");
                    wordDocument.CreateBookmarksForImage(path, word, count);
                }
            }


            ViewBag.Messages = wordDocument.Messages;
            ViewBag.LongText = wordDocument.GetTextFromDocument();
            ViewBag.Footnotes = wordDocument.GetFootnotes();
            return View();
        }

        //private async Task<FileStream> SaveImage(IFormFile image, string path)
        //{
        //    var fileStream = new FileStream(path, FileMode.Create);
        //    await image.CopyToAsync(fileStream);
        //    _logger.LogInformation($"Загружено изображение {image.FileName}");
        //    return fileStream;
        //}

        public IActionResult EditLinks()
        {
            var hyperlinks = wordDocument.GetHyperlinks().ToList();
            ViewBag.Hyperlinks = hyperlinks;
            var bookmarks = wordDocument.GetBookmarks().ToList();
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
            var list = wordDocument.GetHyperlinks().ToList();
            Field field = list[index];
            _logger.LogInformation($"Начато изменение гиперссылки {field.FieldText} на {hypertext}");
            wordDocument.EditLinkInHypertext(field, hypertext);
            _logger.LogInformation($"Завершилось изменение гиперссылки {field.FieldText} на {hypertext}");
            return Redirect("EditLinks");
        }

        public IActionResult DeleteHyperlink(int index)
        {
            var list = wordDocument.GetHyperlinks().ToList();
            Field field = list[index];
            _logger.LogInformation($"Начато удаление гиперссылки {field.FieldText}");
            this.wordDocument.DeleteHyperlink(field);
            _logger.LogInformation($"Завершилось удаление гиперссылки {field.FieldText}");
            return Redirect("EditLinks");
        }

        public IActionResult UpdateBookmark(string bookmark, string text)
        {
            _logger.LogInformation($"Начато изменение закладки {bookmark} на {text}");
            wordDocument.EditTextInBookmark(bookmark, text);
            _logger.LogInformation($"Завершилось изменение закладки {bookmark} на {text}");
            return Redirect("EditLinks");
        }

        public IActionResult DeleteBookmark(string bookmark)
        {
            _logger.LogInformation($"Начато удаление закладки {bookmark}");
            wordDocument.DeleteBookmark(bookmark);
            _logger.LogInformation($"Завершилось удаление закладки {bookmark}");
            return Redirect("EditLinks");
        }

        public FileResult Download(FileFormat fileFormat)
        {
            FolderWork.SetParametresForDownloadFile(fileFormat, userFolder: HomeController.userFolder, filepath: HomeController.filepath, out byte[] fileBytes, out string fileExtension, out string filename);
            _logger.LogInformation($"Начато скачивание файла {filename}");
            return File(fileBytes, "application/" + fileExtension, filename);
        }

        //private static void SetParametresForDownloadFile(FileFormat fileFormat, out byte[] fileBytes, out string fileExtension, out string filename)
        //{
        //    string filepath = new DirectoryInfo(HomeController.userFolder).GetFiles().Where(f => f.FullName == HomeController.filepath).FirstOrDefault().FullName;
        //    switch (fileFormat)
        //    {
        //        case FileFormat.TXT:
        //            filepath = Conversion.ConvertToTxt(filepath);
        //            break;
        //        case FileFormat.RTF:
        //            filepath = Conversion.ConvertToRtf(filepath);
        //            break;
        //        case FileFormat.DOC:
        //            filepath = Conversion.ConvertToWordDoc(filepath);
        //            break;
        //        case FileFormat.DOCX:
        //            filepath = Conversion.ConvertToWordDocx(filepath);
        //            break;
        //        case FileFormat.HTML:
        //            filepath = Conversion.ConvertToHtml(filepath);
        //            break;
        //        case FileFormat.PDF:
        //            filepath = Conversion.ConvertToPdf(filepath);
        //            break;
        //        case FileFormat.ODT:
        //            filepath = Conversion.ConvertToOdt(filepath);
        //            break;
        //        default:
        //            break;
        //    }


        //    FileInfo file = new DirectoryInfo(HomeController.userFolder).GetFiles().Where(f => f.FullName == filepath).FirstOrDefault();
        //    fileBytes = System.IO.File.ReadAllBytes(file.FullName);
        //    fileExtension = file.Extension;
        //    filename = file.Name;
        //}

        private void DeleteFile(FileInfo file)
        {
            file.Delete();
        }

        public IActionResult CreateHyperlinkForExistingImage(int index, string hypertext)
        {
            var image = wordDocument.GetImages().ToList()[index];
            _logger.LogInformation($"Начато создание гиперссылки для существующего изображения с параметрами {image}, {hypertext}");
            wordDocument.CreateHyperlinkForImage(image, hypertext);
            _logger.LogInformation($"Завершилось создание гиперссылки для существующего изображения с параметрами {image}, {hypertext}");
            return Redirect("EditLinks");
        }
    }
}