using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Word.Controllers
{
    public class ReviewController : Controller
    {
        // GET: FileUpload
        public ActionResult Index(string fileName)
        {
            // 伺服器檔案路徑
            string filePath = Server.MapPath("~/FileUpload/Word/" + fileName);

            if (System.IO.File.Exists(filePath))
            {
                // 取得檔案的 MIME 類型
                string mimeType = MimeMapping.GetMimeMapping(filePath);

                // 回傳檔案作為串流，用於瀏覽器預覽
                FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                return new FileStreamResult(fileStream, mimeType);
            }
            else
            {
                return HttpNotFound(); // 若檔案不存在，回傳 404 Not Found
            }
        }
    }
}