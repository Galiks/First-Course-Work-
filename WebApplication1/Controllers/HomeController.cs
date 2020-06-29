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
        readonly IWebHostEnvironment _appEnvironment;
        private readonly ILogger<HomeController> _logger;
        private readonly List<string> formats;
        public static string FileName;
        public static string pathToFile;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment appEnvironment)
        {
            _logger = logger;
            _appEnvironment = appEnvironment;
            formats = new List<string>() { ".docx", ".pdf", ".doc" };
        }

        public async Task<IActionResult> Index(IFormFile file)
        {
            if (file != null)
            {
                var extension = Path.GetExtension(file.FileName);
                if (formats.Contains(extension))
                {
                    string path = _appEnvironment.WebRootPath + @"\Files\Doc\" + file.FileName;
                    pathToFile = path;
                    //FileMode.Append
                    await SaveFile(file, file.FileName, path);
                    return Redirect("Document/Index");
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
                using (var fileStream = new FileStream(path, FileMode.CreateNew))
                {
                    await file.CopyToAsync(fileStream);
                }
            }
            catch (IOException)
            {

                filename = "NEW_" + filename;
                path = _appEnvironment.WebRootPath + @"\Files\Doc\" + filename;
                pathToFile = path;
                FileName = filename;
                await SaveFile(file, filename, path);
            }
        }

        public async Task<IActionResult> LogIn()
        {
            return View();
        }

        public async Task<IActionResult> Registration()
        {
            return View();
        }
    }
}
