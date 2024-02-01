using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Word;

namespace Word.Helper
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

    public class DocxHelper
    {
        public (string, string) MicrosoftOfficeConvertHTML(string rootPath, string fileName, Stream InputStream)
        {
            string path = rootPath + "/FileUpload/MicrosoftOffice";

            var filePath = Path.Combine(path, fileName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            using (var stream = InputStream)
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
                StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
                doc.MainDocumentPart.Document.Body.Remove();
                var tempBody = new DocumentFormat.OpenXml.Wordprocessing.Body(xml);
                var table = tempBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();

                var headerList = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().FirstOrDefault()
                    ?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Select(a => a.InnerText.Trim()).ToList();
                var qusetionIdx = headerList.FindIndex(a => a.Contains("題目 內容") || a.Contains("題目內容"));
                // 跳過首行
                var rows = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Skip(1).ToList();
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
                DocumentFormat.OpenXml.Wordprocessing.TableRow headerRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                var otherTitleCell = headerList.Where(a => !(a.Contains("題目 內容") || a.Contains("題目內容"))).Select(a => new TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new Text(a))))).ToList();
                foreach (var titleCell in otherTitleCell)
                {
                    headerRow.AppendChild(titleCell);
                }
                headerRow.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new Text("題目")))));
                headerRow.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new Text("答案")))));
                headerRow.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new Text("解析")))));
                newTable.AppendChild(headerRow);

                foreach (var row in rows)
                {
                    var cols = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList();
                    Question question = new Question();
                    var bodyRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                    for (int i = 0; i < cols.Count; i++)
                    {
                        var col = cols[i];
                        if (i == qusetionIdx)
                        {
                            question = this.QuestionParse(col, stylePart);
                        }
                        else
                        {
                            var clonedElement = col.CloneNode(true);
                            bodyRow.AppendChild(clonedElement);
                        }
                    }
                    // 標題
                    var titleCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    foreach (var el in question.Title.Elements)
                    {
                        var clonedElement = (OpenXmlElement)el.CloneNode(true);
                        titleCell.AppendChild(clonedElement);
                    }
                    // TableCell至少有一個元素，不然檔案會損壞
                    if (!question.Title.Elements.Any())
                    {
                        titleCell.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                    }
                    bodyRow.AppendChild(titleCell);

                    // 答案
                    var answerCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    foreach (var el in question.Answer.Elements)
                    {
                        var clonedElement = (OpenXmlElement)el.CloneNode(true);
                        answerCell.AppendChild(clonedElement);
                    }
                    // TableCell至少有一個元素，不然檔案會損壞
                    if (!question.Answer.Elements.Any())
                    {
                        answerCell.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                    }
                    bodyRow.AppendChild(answerCell);

                    // 解析
                    var analyzeCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    foreach (var el in question.Analyze.Elements)
                    {
                        var clonedElement = (OpenXmlElement)el.CloneNode(true);
                        analyzeCell.AppendChild(clonedElement);
                    }
                    // TableCell至少有一個元素，不然檔案會損壞
                    if (!question.Analyze.Elements.Any())
                    {
                        analyzeCell.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                    }
                    bodyRow.AppendChild(analyzeCell);
                    newTable.AppendChild(bodyRow);
                }
                newBody.AppendChild(newTable);
                doc.MainDocumentPart.Document.Append(newBody);
                doc.Save();
            }
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            var htmlFileName = fileName.Split('.')[0] + ".html";
            Microsoft.Office.Interop.Word.Document officeDoc = wordApp.Documents.Open(filePath, Encoding.UTF8);
            var htmlPath = Path.Combine(path, htmlFileName);
            officeDoc.SaveAs2(htmlPath, WdSaveFormat.wdFormatFilteredHTML, Encoding.UTF8);
            // 關閉 Word 文件
            officeDoc.Close();
            // 關閉 Word 應用程式
            wordApp.Quit();
            var fontFileName = "BpmfZihiKaiStd-Regular.ttf";
            string fontPath = rootPath + "/Font";
            var fontFullPath = Path.Combine(fontPath, fontFileName);
            HtmlAddFont(htmlPath, "ㄅ字嗨注音標楷 Regular", fontFullPath);
            var htmlUrlPath = Path.Combine("/FileUpload/MicrosoftOffice", htmlFileName);
            var wordUrlPath = Path.Combine("/FileUpload/MicrosoftOffice", fileName);
            return (htmlUrlPath, wordUrlPath);
        }

        private Question QuestionParse(DocumentFormat.OpenXml.Wordprocessing.TableCell cell, StyleDefinitionsPart stylePart)
        {
            var elements = cell.Elements();
            var question = new Question();
            var currentElementType = ElementType.Title;
            string color = null;
            foreach (var element in elements)
            {
                string innerText = element.InnerText.Trim();
                if (innerText.Contains("答案"))
                {
                    string fontColor = FindFontColor(element, color, stylePart);
                    if (fontColor != color)
                    {
                        currentElementType = ElementType.Answer;
                        color = fontColor;
                    }
                }
                else if (innerText.Contains("解析"))
                {
                    string fontColor = FindFontColor(element, color, stylePart);
                    if (fontColor != color)
                    {
                        currentElementType = ElementType.Analyze;
                        color = fontColor;
                    }
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

        private string FindFontColor(OpenXmlElement element, string color, StyleDefinitionsPart stylePart)
        {
            if (element is Run run)
            {
                var runProperties = run.RunProperties;
                var styleId = run.RunProperties?.RunStyle?.Val;
                if (runProperties?.Color?.Val != color)
                {
                    return runProperties?.Color?.Val;
                }
                // 如果抓樣式表以設定的顏色，而不是設置run的
                else if (styleId != null)
                {
                    var runStyle = stylePart.Styles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().FirstOrDefault(s => s.StyleId == styleId);
                    string colorValue = runStyle?.StyleRunProperties?.Color?.Val;
                    if (colorValue != color)
                    {
                        return colorValue;
                    }
                }
            }
            foreach (var el in element.ChildElements)
            {
                var elColor = FindFontColor(el, color, stylePart);
                if (elColor != color)
                {
                    return elColor;
                }
            }
            return color;
        }

        private void HtmlAddFont(string path, string fontName, string fontPath)
        {
            // 建立 HtmlDocument 實例
            HtmlDocument htmlDoc = new HtmlDocument();

            // 載入 HTML 檔案
            htmlDoc.Load(path, Encoding.UTF8);
            // 取得 <html> 標籤，如果不存在，則創建一個並添加 <head> 標籤
            HtmlNode htmlNode = htmlDoc.DocumentNode.SelectSingleNode("//html");
            var hasHtmlNode = htmlNode != null;
            if (!hasHtmlNode)
            {
                htmlNode = HtmlNode.CreateNode("<html></html>");
            }
            // 取得 <head> 標籤，如果不存在，則創建一個並添加到 <html> 標籤中
            HtmlNode headNode = htmlDoc.DocumentNode.SelectSingleNode("//head");
            if (headNode == null)
            {
                headNode = HtmlNode.CreateNode("<head></head>");
            }
            // 將 <head> 標籤加入到 <html> 標籤中
            htmlNode.AppendChild(headNode);
            // 創建新的 <style> 標籤
            HtmlNode styleNode = htmlDoc.CreateElement("style");
            // 讀取字型檔案的二進位資料
            byte[] fontBytes = System.IO.File.ReadAllBytes(fontPath);
            // 將字型檔案二進位資料轉換成 base64 字串
            string base64String = Convert.ToBase64String(fontBytes);
            styleNode.InnerHtml = $"@font-face {{ font-family: '{fontName}'; src: url('data:font/ttf;base64,{base64String}') format('truetype'); }}";

            // 將 <style> 標籤插入到 <head> 標籤中
            headNode.AppendChild(styleNode);
            if (!hasHtmlNode)
            {
                // 加入 <html> 標籤到文檔中
                htmlDoc.DocumentNode.PrependChild(htmlNode);
            }
            htmlDoc.Save(path);
        }
    }
}