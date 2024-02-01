using Core.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Word.Helper;

namespace Core.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly DocxHelper _docxHelper;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(
            ILogger<HomeController> logger,
            DocxHelper docxHelper,
            IWebHostEnvironment hostingEnvironment
        )
        {
            _logger = logger;
            _docxHelper = docxHelper;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public IActionResult Upload(List<IFormFile> files)
        {
            try
            {
                if (files != null && files.Count > 0)
                {
                    string webRootPath = _hostingEnvironment.WebRootPath;
                    (string htmlUrl, string wordUrl) = _docxHelper.MicrosoftOfficeConvertHTML(webRootPath, files[0].FileName, files[0].OpenReadStream());
                    return Json(new { Success = true, HtmlUrl = htmlUrl, WordUrl = wordUrl });
                }
                return Json(new { Success = false });
            }
            catch (Exception ex)
            {
                return Json(new { Success = false });
            }
        }
    }
}