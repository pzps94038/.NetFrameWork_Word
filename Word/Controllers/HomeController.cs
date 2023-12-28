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
using Aspose.Html.Rendering.Doc;

namespace Word.Controllers
{
    public class Question
    {
        public WordElement Title { get; set; } = new WordElement();
        public WordElement Answer { get; set; } = new WordElement();
        public WordElement Analyze { get; set; } = new WordElement();
    }

    public class WordElement 
    {
        public List<OpenXmlElement> Elements { get; set; } = new List<OpenXmlElement> ();
    }

    public enum ElementType
    {
        Title,
        Answer,
        Analyze
    }

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
        public void Upload()
        {
            try
            {
                this.SpireDocConvertHTML(Request.Files[0]);
                //this.AsposeWordsConvertHTML(Request.Files[0]);
            }
            catch (Exception ex)
            {
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
                var xml = doc.MainDocumentPart.Document.Body.OuterXml;
                doc.MainDocumentPart.Document.Body.Remove();
                var tempBody = new DocumentFormat.OpenXml.Wordprocessing.Body(xml);
                var table = tempBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();
                
                var headerList = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().FirstOrDefault()
                    ?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Select(a=> a.InnerText.Trim()).ToList();
                var qusetionIdx = headerList.FindIndex(a => a.Contains("題目 內容") || a.Contains("題目內容"));
                // 跳過首行
                var rows = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Skip(1).Take(1).ToList();
                Question question = new Question();
                foreach (var row in rows) 
                {
                    var cols = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList();
                    for (int i = 0; i < cols.Count; i++) 
                    {
                        var col = cols[i];
                        if (i == qusetionIdx)
                        {
                            question = this.QuestionParse(col);
                        }
                        var colXml = col.OuterXml;
                        var colText = col.InnerText;
                    }
                }
                var newBody = new DocumentFormat.OpenXml.Wordprocessing.Body();
                var newTable = new DocumentFormat.OpenXml.Wordprocessing.Table();
                // 创建表格边框样式和样式特征（这里是示例，你可以根据需求修改样式）
                TableBorders borders = new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = 12 },
                    new BottomBorder() { Val = BorderValues.Single, Size = 12 },
                    new LeftBorder() { Val = BorderValues.Single, Size = 12 },
                    new RightBorder() { Val = BorderValues.Single, Size = 12 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 12 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Size = 12 }
                );

                DocumentFormat.OpenXml.Wordprocessing.TableStyle tableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle() { Val = "TableGrid" }; // 表格样式

                TableProperties tableProperties = new TableProperties();
                tableProperties.Append(borders, tableStyle);

                newTable.AppendChild(tableProperties);

                DocumentFormat.OpenXml.Wordprocessing.TableRow headerRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(new Paragraph(new Run(new Text("題目")))),
                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(new Paragraph(new Run(new Text("答案")))),
                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(new Paragraph(new Run(new Text("解析"))))
                );
                newTable.AppendChild(headerRow);

                var bodyRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

                // 標題
                var titleCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                foreach (var el in question.Title.Elements)
                {
                    var clonedElement = (OpenXmlElement)el.CloneNode(true);
                    titleCell.Append(clonedElement);
                }
                bodyRow.AppendChild(titleCell);

                // 答案
                var answerCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                foreach (var el in question.Answer.Elements)
                {
                    var clonedElement = (OpenXmlElement)el.CloneNode(true);
                    answerCell.Append(clonedElement);
                }
                bodyRow.AppendChild(answerCell);

                // 解析
                var analyzeCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                foreach (var el in question.Analyze.Elements)
                {
                    var clonedElement = (OpenXmlElement)el.CloneNode(true);
                    analyzeCell.Append(clonedElement);
                }
                bodyRow.AppendChild(analyzeCell);
                newTable.AppendChild(bodyRow);
                newBody.Append(newTable);
                doc.MainDocumentPart.Document.Append(newBody);
                doc.Save();
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

        private Question QuestionParse(DocumentFormat.OpenXml.Wordprocessing.TableCell cell) 
        {
            var elements = cell.Elements();
            var question = new Question();
            var currentElementType = ElementType.Title;
            foreach (var element in elements) 
            {
                string innerText = element.InnerText.Trim();

                if (innerText.Contains("答案") && FindFontColor(element, "0000FF"))
                {
                    currentElementType = ElementType.Answer;
                }
                else if (innerText.Contains("解析") && FindFontColor(element, "008000"))
                {
                    currentElementType = ElementType.Analyze;
                }
                // 表格Cell樣式不保留
                if (element is DocumentFormat.OpenXml.Wordprocessing.TableCellProperties) 
                {
                    continue;
                }
                OpenXmlElement clonedElement = (OpenXmlElement)element.CloneNode(true);
                switch (currentElementType)
                {
                    case ElementType.Title:
                        question.Title.Elements.Add(clonedElement);
                        break;
                    case ElementType.Answer:
                        question.Answer.Elements.Add(clonedElement);
                        break;
                    case ElementType.Analyze:
                        question.Analyze.Elements.Add(clonedElement);
                        break;
                }
            }
            return question;
        }

        private bool FindFontColor(OpenXmlElement element, string color)
        {
            if (element is Run run)
            {
                var runProperties = run.RunProperties;
                if (runProperties?.Color?.Val != null && runProperties.Color.Val == color)
                {
                    return true;
                }
            }
            foreach (var el in element.ChildElements)
            {
                if (FindFontColor(el, color))
                {
                    return true;
                }
            }
            return false;
        }
    }
}