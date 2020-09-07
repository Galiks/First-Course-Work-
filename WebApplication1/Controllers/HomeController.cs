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
        public IActionResult Index(IFormFile file)
        {
            if (file != null)
            {
                var extension = Path.GetExtension(file.FileName);

                if (formats.Contains(extension))
                {
                    filepath = userFolder + file.FileName;

                    //await SaveFile(file, file.FileName, filepath);
                    FolderWork.SaveFile(file, file.FileName, ref filepath, ref userFolder);
                    _logger.LogInformation($"Сохранён файл {filepath}");

                    filepath = Conversion.ConvertToWordDocx(filepath);
                    _logger.LogInformation($"Сохранён переконвентированный файл {filepath}");

                    if (string.IsNullOrWhiteSpace(filepath))
                    {
                        _logger.LogInformation($"Преобразование файла {filepath} невозможно");
                        ViewBag.FileFormatErrorMessage = "Неверный формат файла. Дальнейшее преобразование невозможно!";
                        return View();
                    }

                    return RedirectToAction("Index", "Document");
                }
                else
                {
                    _logger.LogInformation($"Введён не тот формат файла - {extension}. Файл - {file.FileName}.");
                    ViewBag.FileFormatErrorMessage = "Неверный формат файла. Должен быть TXT, RTF, HTML, ODT, DOC, DOCX, PDF";
                    return View();
                }
            }
            return View();
        }

        //private async Task SaveFile(IFormFile file, string filename, string path)
        //{
        //    if (System.IO.File.Exists(path))
        //    {
        //        filename = "NEW_" + filename;
        //        filepath = userFolder + filename;
        //        await SaveFile(file, filename, filepath);
        //    }
        //    else
        //    {
        //        using var fileStream = new FileStream(path, FileMode.CreateNew);
        //        await file.CopyToAsync(fileStream);
        //    }
        //}

        public IActionResult LogIn(string lastName, string firstName, string patronymic)
        {
            if (string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(patronymic))
            {
                return View();
            }
            else
            {
                string initials = FolderWork.GetFolderName(firstName, lastName, patronymic);
                string absolutPath = _appEnvironment.WebRootPath + @"\Files";
                try
                {
                    foreach (var item in Directory.GetDirectories(absolutPath))
                    {
                        if (item.Contains(initials))
                        {
                            userFolder = item + @"\";
                        }
                    }
                }
                catch (System.Exception)
                {
                    _logger.LogError($"Абсолютный путь {absolutPath} не найден.");
                    Directory.CreateDirectory(absolutPath);
                    _logger.LogInformation($"Создан абсолютный путь {absolutPath}.");
                    
                }


                if (!string.IsNullOrWhiteSpace(userFolder))
                {
                    _logger.LogInformation($"Пользователь {firstName} {lastName} {patronymic} вошёл в систему.");
                    return RedirectToAction("Index");
                }
                else
                {
                    _logger.LogError($"Пользователь {firstName} {lastName} {patronymic} не был найден.");
                    return View();
                }
            }

        }

        public IActionResult Registration(string lastName, string firstName, string patronymic)
        {
            if (string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(patronymic))
            {
                _logger.LogError($"Поля при регистрации оказались пустыми пустые.");
                return View();
            }
            else
            {
                userFolder = FolderWork.CreateFolder(firstName, lastName, patronymic, _appEnvironment.WebRootPath + @"\Files") + @"\";
                if (!string.IsNullOrWhiteSpace(userFolder))
                {
                    _logger.LogInformation($"Папка для пользователя {firstName} {lastName} {patronymic} была успешно создана.");
                    return RedirectToAction("Index"); 
                }
                else
                {
                    _logger.LogError($"Папка для пользователя {firstName} {lastName} {patronymic} не была создана.");
                    return View();
                }
            }
        }
    }
}
