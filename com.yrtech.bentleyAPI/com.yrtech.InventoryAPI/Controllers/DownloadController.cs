using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace com.yrtech.InventoryAPI.Controllers
{
    [RoutePrefix("download")]
    public class DownloadController : Controller
    {
        [HttpGet]
        [Route("DownloadFile")]
        public void DownloadFile(string filePath, bool isDeleteAfterDownload = false)
        {
            if (!string.IsNullOrWhiteSpace(filePath) && System.IO.File.Exists(filePath))
            {
                string filename = Path.GetFileName(filePath);
                FileStream stream = null;
                using (stream = new FileStream(filePath, FileMode.Open))
                {
                    byte[] bytes = new byte[(int)stream.Length];
                    stream.Position = 0;
                    stream.Read(bytes, 0, bytes.Length);
                    stream.Close();
                    Response.Clear();
                    Response.Charset = "UTF-8";
                    Response.ContentEncoding = Encoding.GetEncoding("UTF-8");
                    Response.AddHeader("content-type", "application/x-msdownload");
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(filename, Encoding.GetEncoding("UTF-8")));
                    Response.BinaryWrite(bytes);
                    Response.End();
                    if (isDeleteAfterDownload)
                    {
                        System.IO.File.Delete(filePath);
                    }
                }
            }
        }
    }
}