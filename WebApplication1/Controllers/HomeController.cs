using BusinessLogicLayer;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        readonly IWebHostEnvironment _appEnvironment;
        private readonly ILogger<HomeController> _logger;
        private readonly List<string> formats;

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
                    string path = _appEnvironment.WebRootPath + "/Files/Doc/" + file.FileName;
                    using (var fileStream = new FileStream(path, FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }
                }
                else
                {
                    ViewBag.FileFormatErrorMessage = "Неверный формат файла. Должен быть DOC, PDF, DOCX";
                }
            }
            return View();
        }
    }
}
