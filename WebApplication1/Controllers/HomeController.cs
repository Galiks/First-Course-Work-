using BusinessLogicLayer;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private readonly string absolutPath;
        readonly IWebHostEnvironment _appEnvironment;
        private readonly ILogger<HomeController> _logger;
        private readonly List<string> formats;
        public static string FileName;
        public static string pathToFile;
        public static string userFolder;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment appEnvironment)
        {
            _logger = logger;
            _appEnvironment = appEnvironment;
            absolutPath = _appEnvironment.WebRootPath + @"\Files";
            formats = new List<string>() { ".docx", ".pdf", ".doc" };
            
        }

        public async Task<IActionResult> Index(IFormFile file)
        {
            if (file != null)
            {
                var extension = Path.GetExtension(file.FileName);
                if (formats.Contains(extension))
                {
                    string path = userFolder + file.FileName;
                    pathToFile = path;
                    //FileMode.Append
                    await SaveFile(file, file.FileName, path);
                    return RedirectToAction("Index", "Document");
                }
                else
                {
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
                FileName = filename;
                using var fileStream = new FileStream(path, FileMode.CreateNew);
                await file.CopyToAsync(fileStream);
            }
            catch (IOException)
            {

                filename = "NEW_" + filename;
                path = userFolder + @"\" + filename;
                pathToFile = path;
                FileName = filename;
                await SaveFile(file, filename, path);
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
                var list = Directory.GetDirectories(absolutPath);
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

        public IActionResult Registration(string lastName, string firstName, string patronymic, string password)
        {
            if (string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(password) || string.IsNullOrWhiteSpace(patronymic))
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
