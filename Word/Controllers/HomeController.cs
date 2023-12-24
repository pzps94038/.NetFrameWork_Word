using Aspose.Html;
using Aspose.Words;
using Aspose.Words.Saving;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GemBox.Document;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Web;
using System.Web.Mvc;
using Document = Spire.Doc.Document;
using Mammoth;
using SautinSoft;
using Spire.Doc.Interface;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Aspose.Html.Forms;
using DocumentFormat.OpenXml;
using Syncfusion.DocIO.DLS;
using System.Runtime.InteropServices;

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
                this.SpireDocConvertHTML(Request.Files[0]);
                //this.AsposeWordsConvertHTML(Request.Files[0]);
                

                return Json(new
                {
                    path = ""
                });
            }
            catch (Exception ex)
            {
                return Json(ex.Message);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// https://purchase.aspose.com/policies/license-types/
        /// <param name="file"></param>
        private void AsposeWordsConvertHTML(HttpPostedFileBase file) 
        {
            string path = Server.MapPath("~/FileUpload/AsposeWords");
            var fileName = file.FileName;

            var filePath = Path.Combine(path, fileName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            using (var stream = file.InputStream)
            {
                var buffer = new byte[stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                System.IO.File.WriteAllBytes(filePath, buffer);
            }

            // 文件複製
            // System.IO.File.Copy(origPath, filePath);
            // 文件處理
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
            }
            var htmlFileName = fileName.Split('.')[0] + ".html";
            Aspose.Words.Document asposeDoc = new Aspose.Words.Document(filePath);
            // 設置 HTML 轉換選項
            Aspose.Words.Saving.HtmlSaveOptions options = new Aspose.Words.Saving.HtmlSaveOptions();
            options.SaveFormat = SaveFormat.Html;
            options.ExportImagesAsBase64 = true;
            options.CssStyleSheetType = Aspose.Words.Saving.CssStyleSheetType.Inline;
            string fontPath = Server.MapPath("~/Font");
            // 設置自訂字體
            options.FontsFolder = fontPath; // 自訂字體文件夾路徑
            asposeDoc.Save(Path.Combine(path, htmlFileName), options);
        }

        /// <summary>
        /// https://www.e-iceblue.com/Buy/Spire.Doc.html
        /// </summary>
        /// <param name="file"></param>
        private void SpireDocConvertHTML(HttpPostedFileBase file)
        {
            string path = Server.MapPath("~/FileUpload/Spire.Doc");
            var fileName = file.FileName;

            var filePath = Path.Combine(path, fileName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            using (var stream = file.InputStream)
            {
                var buffer = new byte[stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                System.IO.File.WriteAllBytes(filePath, buffer);
            }

            // 文件複製
            // System.IO.File.Copy(origPath, filePath);
            // 文件處理
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            { 
            }
            //Create a Document instance
            Document spireDoc = new Spire.Doc.Document(filePath);
            spireDoc.HtmlExportOptions.ImageEmbedded = true;
            spireDoc.HtmlExportOptions.CssStyleSheetType = Spire.Doc.CssStyleSheetType.Internal;
            var fontFileName = "BpmfZihiKaiStd-Regular.ttf";
            string fontPath = Server.MapPath("~/Font");
            var fontFullPath = Path.Combine(fontPath, fontFileName);
            spireDoc.PrivateFontList.Add(new PrivateFontPath("BpmfZihiKaiStd-Regular", fontFullPath));
            var htmlFileName = fileName.Split('.')[0] + ".html";
            spireDoc.SaveToFile(Path.Combine(path, htmlFileName), FileFormat.Html);
        }

        public static void GetLasChild(List<OpenXmlElement> openXml) 
        {
            foreach (var el in openXml) 
            {
                if (el.HasChildren)
                {
                    GetLasChild(el.ChildElements.ToList());
                }
                else 
                {
                    Console.Write("ao6bk3");
                }
            }
        }
    }


}