using BusinessLogicLayer;
using BusinessLogicLayer.Conversion;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;

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
            try
            {
                formats = WordDocument.GetFileFormats();
            }
            catch (Exception e)
            {
                WriteExceptionInLog(e);
            }
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

                    try
                    {
                        filepath = Conversion.ConvertToWordDocx(filepath);
                    }
                    catch (Exception e)
                    {
                        WriteExceptionInLog(e);
                        return View();
                    }
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
                string initials = null;
                try
                {
                    initials = FolderWork.GetFolderName(firstName, lastName, patronymic);
                }
                catch (Exception e)
                {
                    WriteExceptionInLog(e);
                    return View();
                }

                if (string.IsNullOrWhiteSpace(initials))
                {
                    _logger.LogError($"Инициалы пользователя пусты.");
                }

                string absolutPath = _appEnvironment.WebRootPath + @"\Files";            

                if (!Directory.Exists(absolutPath))
                {
                    _logger.LogError($"Абсолютный путь {absolutPath} не найден.");
                    Directory.CreateDirectory(absolutPath);
                    _logger.LogInformation($"Создан абсолютный путь {absolutPath}.");
                    return View();
                }
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
                catch (Exception e)
                {
                    _logger.LogError($"Абсолютный путь {absolutPath} не найден.");
                    WriteExceptionInLog(e);
                    Directory.CreateDirectory(absolutPath);
                    _logger.LogInformation($"Создан абсолютный путь {absolutPath}.");
                    return View();
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
                try
                {
                    userFolder = FolderWork.CreateFolder(firstName, lastName, patronymic, _appEnvironment.WebRootPath + @"\Files") + @"\";
                }
                catch (Exception e)
                {
                    WriteExceptionInLog(e);
                    return View();
                }
                if (!string.IsNullOrWhiteSpace(userFolder))
                {
                    _logger.LogInformation($"Папка для пользователя {firstName} {lastName} {patronymic} была успешно создана.");
                    return RedirectToAction("Login");
                }
                else
                {
                    _logger.LogError($"Папка для пользователя {firstName} {lastName} {patronymic} не была создана.");
                    return View();
                }
            }
        }

        private void WriteExceptionInLog(Exception e)
        {
            _logger.LogError($"Message {e.Message} {(e?.InnerException is null ? "" : string.Join(Environment.NewLine, "InnerException: ", e.InnerException.Message))}");
        }
    }
}
