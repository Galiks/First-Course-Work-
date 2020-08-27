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
using BusinessLogicLayer.Conversion;

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
        [RequestSizeLimit(10000000)]
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

                    filepath = Conversion.ConvertToWordDocx(filepath);

                    if (string.IsNullOrWhiteSpace(filepath))
                    {
                        _logger.LogInformation("Преобразование невозможно");
                        ViewBag.FileFormatErrorMessage = "Неверный формат файла. Дальнейшее преобразование невозможно!";
                        return View();
                    }

                    return RedirectToAction("Index", "Document");
                }
                else
                {
                    _logger.LogInformation("Введён не тот формат.");
                    ViewBag.FileFormatErrorMessage = "Неверный формат файла. Должен быть TXT, RTF, HTML, ODT, DOC, DOCX, PDF";
                    return View();
                }
            }
            return View();
        }

        private async Task SaveFile(IFormFile file, string filename, string path)
        {
            if (System.IO.File.Exists(path))
            {
                //filepath = userFolder + @"\" + "NEW_" + filename;
                filename = "NEW_" + filename;
                filepath = userFolder + filename;
                await SaveFile(file, filename, filepath);
            }
            else
            {
                using var fileStream = new FileStream(path, FileMode.CreateNew);
                await file.CopyToAsync(fileStream);
            }
            //try
            //{
            //    using var fileStream = new FileStream(path, FileMode.CreateNew);
            //    await file.CopyToAsync(fileStream);
            //}
            //catch (IOException)
            //{
            //    filepath = userFolder + @"\" + "NEW_" + filename;
            //    await SaveFile(file, filename, filepath);
            //}
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
