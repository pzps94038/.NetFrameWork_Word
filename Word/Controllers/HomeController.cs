using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Word;
using Word.Helper;

namespace Word.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult Upload()
        {
            try
            {
                var docxHelper = new DocxHelper();
                (string htmlUrl, string wordUrl) = docxHelper.MicrosoftOfficeConvertHTML(Server.MapPath("~/"), Request.Files[0].FileName, Request.Files[0].InputStream);
                // this.SpireDocConvertHTML(Request.Files[0]);
                // this.AsposeWordsConvertHTML(Request.Files[0]);
                return Json(new { Success = true, HtmlUrl = htmlUrl, WordUrl = wordUrl });
            }
            catch (Exception ex)
            {
                return Json(new { Success = false });
            }
        }
    }
}