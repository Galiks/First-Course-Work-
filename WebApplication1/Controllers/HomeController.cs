using BusinessLogicLayer;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Pdf;
using Spire.Doc;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Spire.Doc.Documents;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        readonly IWebHostEnvironment _appEnvironment;
        private readonly ILogger<HomeController> _logger;
        private readonly List<string> formats;
        public static string filepath;
        //public static string pathToFile;
        public static string userFolder;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment appEnvironment)
        {
            _logger = logger;
            _logger.LogDebug(1, "NLog injected into HomeController");
            _appEnvironment = appEnvironment;
            formats = WordDocument.GetFileFormats();
        }

        public async Task<IActionResult> Index(IFormFile file)
        {
            if (file != null)
            {
                var extension = Path.GetExtension(file.FileName);

                if (formats.Contains(extension))
                {
                    filepath = userFolder + file.FileName;
                    //pathToFile = path;



                    //FileMode.Append
                    await SaveFile(file, file.FileName, filepath);

                    if (extension.Equals(".pdf"))
                    {
                        PdfDocument pdfDoc = new PdfDocument();
                        pdfDoc.LoadFromFile(filepath);
                        //pathToFile = $"{path}.doc";
                        filepath = $"{filepath}.doc";
                        pdfDoc.SaveToFile(filepath, Spire.Pdf.FileFormat.DOC);
                        //FileName = pathToFile;
                    }
                    else if (extension.Equals(".svg"))
                    {
                        //PdfDocument doc1 = new PdfDocument();
                        //doc1.LoadFromSvg(path);
                    }
                    else if (extension.Equals(".html"))
                    {
                        Document document = new Document();
                        document.LoadFromFile(filepath, Spire.Doc.FileFormat.Html, XHTMLValidationType.None);

                    }
                    else if (extension.Equals(".doc") || extension.Equals(".docx"))
                    {
                        //Document doc = new Document(filepath);
                        //doc.SaveToFile($"{path}.doc", Spire.Doc.FileFormat.Doc);
                    }
                    else
                    {
                        Document doc = new Document(filepath);
                        doc.SaveToFile($"{filepath}.doc", Spire.Doc.FileFormat.Doc);
                    }



                    return RedirectToAction("Index", "Document");
                }
                else
                {
                    _logger.LogInformation("Введён не тот формат.");
                    ViewBag.FileFormatErrorMessage = "Неверный формат файла. Должен быть DOC, PDF, DOCX";
                    return View();
                }
            }
            return View();
        }

        private async Task SaveFile(IFormFile file, string filename, string path)
        {
            try
            {
                using var fileStream = new FileStream(path, FileMode.CreateNew);
                await file.CopyToAsync(fileStream);
            }
            catch (IOException)
            {
                //path = userFolder + @"\" + "NEW_" + filename;
                //pathToFile = path;
                filepath = userFolder + @"\" + "NEW_" + filename; ;
                await SaveFile(file, filename, filepath);
            }
        }

        public IActionResult LogIn(string lastName, string firstName, string patronymic)
        {
            if (string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(patronymic))
            {
                return View();
            }
            else
            {
                //Directory.SetCurrentDirectory(absolutPath);
                string initials = FolderWork.GetFolderName(firstName, lastName, patronymic);
                string absolutPath = _appEnvironment.WebRootPath + @"\Files";
                foreach (var item in Directory.GetDirectories(absolutPath))
                {
                    if (item.Contains(initials))
                    {
                        userFolder = item + @"\";
                    }
                }


                if (!string.IsNullOrWhiteSpace(userFolder))
                {
                    return RedirectToAction("Index");
                }
                else
                {
                    return View();
                }
            }

        }

        public IActionResult Registration(string lastName, string firstName, string patronymic)
        {
            if (string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(patronymic))
            {
                return View();
            }
            else
            {
                userFolder = FolderWork.CreateFolder(firstName, lastName, patronymic, _appEnvironment.WebRootPath + @"\Files") + @"\";
                return RedirectToAction("Index");
            }
        }
    }
}
