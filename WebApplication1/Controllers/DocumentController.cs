using BusinessLogicLayer;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc.Fields;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class DocumentController : Controller
    {
        private readonly IWebHostEnvironment _appEnvironment;
        private readonly ILogger<DocumentController> _logger;
        private readonly WordDocument wordDocument;

        public DocumentController(IWebHostEnvironment appEnvironment, ILogger<DocumentController> logger)
        {
            _appEnvironment = appEnvironment;
            _logger = logger;
            _logger.LogInformation($"В класс WordDocument передан файл {HomeController.filepath}.");
            wordDocument = new WordDocument(HomeController.filepath);
        }

        public async Task<IActionResult> Index(string word, string text, string linkType, IFormFile image, string oneOrAll, int count = default)
        {
            try
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
                    catch (Exception e)
                    {
                        _logger.LogError($"Не удалось сохранить изображение {image.FileName}.");
                        WriteExceptionInLog(e);
                        ViewBag.Messages = e.Message.ToString();
                        ViewBag.LongText = wordDocument.GetTextFromDocument();
                        return View();
                    }
                }

                if (!string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(linkType))
                {
                    if (linkType.Equals("bookmark"))
                    {
                        try
                        {
                            _logger.LogInformation($"Начато создание закладки для текста с параметрами {text}, {word}");
                            if (oneOrAll.Equals("all"))
                            {
                                wordDocument.CreateBookmarksForText(word, text, count);
                            }
                            else if (oneOrAll.Equals("one"))
                            {
                                wordDocument.CreateBookmarkForOneWord(word, text);
                            }
                            _logger.LogInformation($"Завершилось создание закладки для текста с параметрами {text}, {word}");
                        }
                        catch (Exception e)
                        {
                            _logger.LogError($"Message {e.Message}");
                            WriteExceptionInLog(e);
                            ViewBag.Messages = e.Message.ToString();
                            ViewBag.LongText = wordDocument.GetTextFromDocument();
                            return View();
                        }
                    }
                    else if (linkType.Equals("hyperlink"))
                    {
                        try
                        {
                            _logger.LogInformation($"Начато создание гиперссылки для текста с параметрами {text}, {word}");
                            if (oneOrAll.Equals("all"))
                            {
                                wordDocument.CreateHyperlinksForText(word, text, count);
                            }
                            else if (oneOrAll.Equals("one"))
                            {
                                wordDocument.CreateHyperlinkForOneWord(word, text);
                            }
                            _logger.LogInformation($"Завершилось создание гиперссылки для текста с параметрами {text}, {word}");
                        }
                        catch (Exception e)
                        {
                            _logger.LogError($"Message {e.Message}");
                            WriteExceptionInLog(e);
                            ViewBag.Messages = e.Message.ToString();
                            ViewBag.LongText = wordDocument.GetTextFromDocument();
                            return View();
                        }
                    }
                }
                else if (!string.IsNullOrWhiteSpace(path) & !string.IsNullOrWhiteSpace(text) & !string.IsNullOrWhiteSpace(linkType))
                {
                    if (linkType.Equals("hyperlink"))
                    {
                        try
                        {
                            _logger.LogInformation($"Начато создание гиперссылки для изображения с параметрами {text}, {path}");
                            wordDocument.CreatHyperlinkForImage(path, text);
                            _logger.LogInformation($"Завершилось создание гиперссылки для изображения с параметрами {text}, {path}");
                        }
                        catch (Exception e)
                        {
                            _logger.LogError($"Message {e.Message}");
                            WriteExceptionInLog(e);
                            ViewBag.Messages = e.Message.ToString();
                            ViewBag.LongText = wordDocument.GetTextFromDocument();
                            return View();
                        }
                    }
                }
                else if (!string.IsNullOrWhiteSpace(path) & !string.IsNullOrWhiteSpace(word) & !string.IsNullOrWhiteSpace(linkType))
                {
                    if (linkType.Equals("bookmark"))
                    {
                        try
                        {
                            _logger.LogInformation($"Начато создание закладки для изображения с параметрами {text}, {path}");
                            if (oneOrAll.Equals("all"))
                            {
                                wordDocument.CreateBookmarksForImage(path, word, count);
                            }
                            else if (oneOrAll.Equals("one"))
                            {
                                wordDocument.CreateBookmarkByImgeForOneWord(path, word);
                            }
                            _logger.LogInformation($"Завершилось создание закладки для изображения с параметрами {text}, {path}");
                        }
                        catch (Exception e)
                        {
                            _logger.LogError($"Message {e.Message}");
                            WriteExceptionInLog(e);
                            ViewBag.Messages = e.Message.ToString();
                            ViewBag.LongText = wordDocument.GetTextFromDocument();
                            return View();
                        }
                    }
                }



                ViewBag.Messages = wordDocument.Messages;
                ViewBag.LongText = wordDocument.GetTextFromDocument();
                return View();
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                ViewBag.LongText = wordDocument.GetTextFromDocument();
                return View();
            }
            
            
        }

        public IActionResult EditLinks()
        {
            try
            {
                var hyperlinks = wordDocument.GetHyperlinks().ToList();
                ViewBag.Hyperlinks = hyperlinks;
                var textBookmarks = wordDocument.GetTextBookmarks().ToList();
                ViewBag.Bookmarks = textBookmarks;
                var imageBookmarks = wordDocument.GetImageBookmarks().ToList();
                ViewBag.ImageBookmarks = imageBookmarks;
                var images = wordDocument.GetImages().ToList();
                ViewBag.Images = images;
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                return View();
            }

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
            try
            {
                _logger.LogInformation($"Начато изменение гиперссылки {field.FieldText} на {hypertext}");
                wordDocument.EditLinkInHypertext(field, hypertext);
                _logger.LogInformation($"Завершилось изменение гиперссылки {field.FieldText} на {hypertext}");
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                return Redirect("EditLinks");
            }
            return Redirect("EditLinks");
        }

        public IActionResult DeleteHyperlink(int index)
        {
            var list = wordDocument.GetHyperlinks().ToList();
            Field field = list[index];
            try
            {
                _logger.LogInformation($"Начато удаление гиперссылки {field.FieldText}");
                wordDocument.DeleteHyperlink(field);
                _logger.LogInformation($"Завершилось удаление гиперссылки {field.FieldText}");
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                return Redirect("EditLinks");
            }
            return Redirect("EditLinks");
        }

        public async Task<IActionResult> UpdateBookmarkAsync(string bookmark, string text, IFormFile image)
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
                catch (Exception e)
                {
                    _logger.LogError($"Не удалось сохранить изображение {image.FileName}.");
                    WriteExceptionInLog(e);
                    ViewBag.Messages = e.Message.ToString();
                    ViewBag.LongText = wordDocument.GetTextFromDocument();
                    return View();
                }
            }

            try
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    _logger.LogInformation($"Начато изменение закладки {bookmark} на {text}");
                    wordDocument.EditTextInBookmark(bookmark, text);
                    _logger.LogInformation($"Завершилось изменение закладки {bookmark} на {text}");
                } 
                else
                {
                    _logger.LogInformation($"Начато изменение изображения закладки {bookmark}");
                    wordDocument.EditImageInBookmark(bookmark, path);
                    _logger.LogInformation($"Завершилось изменение изображения закладки {bookmark}");
                }
                
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                return Redirect("EditLinks");
            }
            return Redirect("EditLinks");
        }

        public IActionResult DeleteBookmark(string bookmark)
        {
            try
            {
                _logger.LogInformation($"Начато удаление закладки {bookmark}");
                wordDocument.DeleteBookmark(bookmark);
                _logger.LogInformation($"Завершилось удаление закладки {bookmark}");
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                return Redirect("EditLinks");
            }
            return Redirect("EditLinks");
        }

        public FileResult Download(FileFormat fileFormat)
        {
            try
            {
                FolderWork.SetParametresForDownloadFile(fileFormat, userFolder: HomeController.userFolder, filepath: HomeController.filepath, out byte[] fileBytes, out string fileExtension, out string filename);
                _logger.LogInformation($"Начато скачивание файла {filename}");
                return File(fileBytes, "application/" + fileExtension, filename);
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                return null;
            }
        }

        private void DeleteFile(FileInfo file)
        {
            file.Delete();
        }

        public IActionResult CreateHyperlinkForExistingImage(int index, string hypertext)
        {
            var image = wordDocument.GetImages().ToList()[index];
            try
            {
                _logger.LogInformation($"Начато создание гиперссылки для существующего изображения с параметрами {image}, {hypertext}");
                wordDocument.CreateHyperlinkForImage(image, hypertext);
                _logger.LogInformation($"Завершилось создание гиперссылки для существующего изображения с параметрами {image}, {hypertext}");
            }
            catch (Exception e)
            {
                _logger.LogError($"Message {e.Message}");
                WriteExceptionInLog(e);
                ViewBag.Messages = e.Message.ToString();
                return Redirect("EditLinks");
            }
            return Redirect("EditLinks");
        }

        private void WriteExceptionInLog(Exception e)
        {
            string innerException = null;
            if (e?.InnerException != null)
            {
                innerException = GetInnerException(e.InnerException);
            }
            _logger.LogError($"Message {e.Message} | {innerException}");
        }

        private string GetInnerException(Exception innerException)
        {
            return $"{Environment.NewLine} InnerException: {innerException.Message}";
        }
    }
}