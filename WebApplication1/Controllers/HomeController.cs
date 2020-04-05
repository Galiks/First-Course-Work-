using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
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
        public static string filename;

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
                    TempData["Path"] = path;
                    filename = path;
                    using (var fileStream = new FileStream(path, FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }
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
    }
}
