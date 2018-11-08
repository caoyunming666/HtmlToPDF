using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Pechkin;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace HtmlToPDF_Demo.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult DownloadPdf2()
        {
            string url = "http://www.baidu.com/";

            using (IPechkin pechkin = Factory.Create(new GlobalConfig()))
            {
                ObjectConfig oc = new ObjectConfig();
                oc.SetPrintBackground(true);
                oc.SetLoadImages(true);
                oc.SetScreenMediaType(true);
                oc.SetPageUri(url);
                byte[] pdf = pechkin.Convert(oc);
                return File(pdf, "application/pdf", "ExportPdf.pdf");
            }
        }

        /// <summary>
        /// 執行此Url，下載PDF檔案
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadPdf()
        {
            string htmlText = GetViewHtml(ControllerContext, "Index", null);

            if (string.IsNullOrEmpty(htmlText))
            {
                return null;
            }

            //string compHtml = "<html><head><style>@font-face{font-family: \"宋体\";src:url(\"C:\\Windows\\Fonts\\simsun.ttf\");}</style></head><body>";

            string compHtml = "<html><head></head><body>";


            //避免當htmlText無任何html tag標籤的純文字時，轉PDF時會掛掉，所以一律加上<p>標籤
            compHtml += "<p>" + htmlText + "</p></body></html>";
            string tempPostContent = getImage(compHtml);

            #region 字符串，字节数组相互转换
            //byte[] data = Encoding.BigEndianUnicode.GetBytes(tempPostContent);//来取得字节
            //string str = Encoding.BigEndianUnicode.GetString(data);//来取得字符.
            //byte[] data = Encoding.UTF8.GetBytes(tempPostContent);//字串轉成byte[]
            //string str = Encoding.Default.GetString(data); 
            #endregion

            HtmlTextConvertToPdf(tempPostContent);

            //byte[] pdfFile = this.ConvertHtmlTextToPDF(htmlText);
            //return new BinaryContentResult(pdfFile, "application/pdf");
            return Json(new { }, JsonRequestBehavior.AllowGet);
        }

        private byte[] ConvertHtmlTextToPDF(string htmlText)
        {
            if (string.IsNullOrEmpty(htmlText))
            {
                return null;
            }
            //避免當htmlText無任何html tag標籤的純文字時，轉PDF時會掛掉，所以一律加上<p>標籤
            htmlText = "<p>" + htmlText + "</p>";

            string tempPostContent = getImage(htmlText);

            MemoryStream outputStream = new MemoryStream();//要把PDF寫到哪個串流
            byte[] data = Encoding.UTF8.GetBytes(tempPostContent);//字串轉成byte[]
            MemoryStream msInput = new MemoryStream(data);
            Document doc = new Document();//要寫PDF的文件，建構子沒填的話預設直式A4
            PdfWriter writer = PdfWriter.GetInstance(doc, outputStream);
            //指定文件預設開檔時的縮放為100%
            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            //開啟Document文件 
            doc.Open();
            //使用XMLWorkerHelper把Html parse到PDF檔裡
            XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, msInput, null, Encoding.UTF8, new UnicodeFontFactory());
            //將pdfDest設定的資料寫到PDF檔
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, writer);
            writer.SetOpenAction(action);
            doc.Close();
            msInput.Close();
            outputStream.Close();
            //回傳PDF檔案 
            return outputStream.ToArray();

        }

        private string getImage(string input)
        {
            if (input == null)
                return string.Empty;
            string tempInput = input;
            string pattern = @"<img(.|\n)+?>";
            string src = string.Empty;

            HttpContextBase context = ControllerContext.HttpContext;

            //Change the relative URL's to absolute URL's for an image, if any in the HTML code.
            foreach (Match m in Regex.Matches(input, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.RightToLeft))
            {
                if (m.Success)
                {
                    string tempM = m.Value;
                    string pattern1 = "src=[\'|\"](.+?)[\'|\"]";
                    Regex reImg = new Regex(pattern1, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                    Match mImg = reImg.Match(m.Value);

                    if (mImg.Success)
                    {
                        src = mImg.Value.ToLower().Replace("src=", "").Replace("\"", "");

                        if (src.ToLower().Contains("http://") == false)
                        {
                            //Insert new URL in img tag
                            src = "src=\"" + context.Request.Url.Scheme + "://" +
                            context.Request.Url.Authority + src + "\"";
                            try
                            {
                                tempM = tempM.Remove(mImg.Index, mImg.Length);
                                tempM = tempM.Insert(mImg.Index, src);

                                //insert new url img tag in whole html code
                                tempInput = tempInput.Remove(m.Index, m.Length);
                                tempInput = tempInput.Insert(m.Index, tempM);
                            }
                            catch (Exception e)
                            {

                            }
                        }
                    }
                }
            }
            return tempInput;
        }

        /// <summary>
        /// 在控制器内获取指定视图生成后的 HTML
        /// </summary>
        /// <param name="context">当前控制器的上下文</param>
        /// <param name="viewName">视图名称</param>
        /// <param name="param">视图所需要的参数</param>
        /// <returns>视图生成的HTML</returns>
        private string GetViewHtml(ControllerContext context, string viewName, Object param)
        {
            if (string.IsNullOrEmpty(viewName))
            {
                viewName = context.RouteData.GetRequiredString("action");
            }
            context.Controller.ViewData.Model = param;
            using (StringWriter sw = new StringWriter())
            {
                ViewEngineResult viewResult = ViewEngines.Engines.FindPartialView(context, viewName);
                ViewContext viewContext = new ViewContext(context,
                                                  viewResult.View,
                                                  context.Controller.ViewData,
                                                  context.Controller.TempData,
                                                  sw);
                try
                {
                    viewResult.View.Render(viewContext, sw);
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("无法获取视图生成的HTML代码");
                }
                return sw.GetStringBuilder().ToString();
            }
        }





        /// <summary>
        /// HTML文本内容转换为PDF
        /// </summary>
        /// <param name="strHtml">HTML文本内容</param>
        /// <param name="savePath">PDF文件保存的路径</param>
        /// <returns></returns>
        public bool HtmlTextConvertToPdf(string strHtml)
        {
            string appRootPath = AppDomain.CurrentDomain.BaseDirectory;
            var fullpath = appRootPath + "/ConvertToPdf";
            if (!Directory.Exists(fullpath))
            {
                Directory.CreateDirectory(fullpath);
            }
            string fileName = fullpath + "/" + DateTime.Now.ToString("yyyyMMddHHmmss") + new Random().Next(1000, 10000) + ".pdf";

            bool flag = false;
            try
            {
                string htmlPath = HtmlTextConvertFile(strHtml);

                flag = HtmlConvertToPdf(htmlPath, fileName);
                //System.IO.File.Delete(htmlPath);
            }
            catch
            {
                flag = false;
            }
            return flag;
        }


        /// <summary>
        /// HTML转换为PDF
        /// </summary>
        /// <param name="htmlPath">可以是本地路径，也可以是网络地址</param>
        /// <param name="savePath">PDF文件保存的路径</param>
        /// <returns></returns>
        public bool HtmlConvertToPdf(string htmlPath, string savePath)
        {
            bool flag = false;
            CheckFilePath(savePath);

            ///这个路径为程序集的目录，因为我把应用程序 wkhtmltopdf.exe 放在了程序集同一个目录下
            string exePath = @"D:\ITtools\wkhtmltopdf\bin\wkhtmltopdf.exe";
            if (!System.IO.File.Exists(exePath))
            {
                throw new Exception("No application wkhtmltopdf.exe was found.");
            }

            try
            {
                /* 
                 * 在程序中调用另一个程序
                 *      调用cmd窗口也是这样写
                 */
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                processStartInfo.FileName = exePath;
                processStartInfo.WorkingDirectory = Path.GetDirectoryName(exePath);
                processStartInfo.UseShellExecute = false;
                processStartInfo.CreateNoWindow = true;
                processStartInfo.RedirectStandardInput = true;
                processStartInfo.RedirectStandardOutput = true;
                processStartInfo.RedirectStandardError = true;
                processStartInfo.Arguments = GetArguments(htmlPath, savePath);

                Process process = new Process();
                process.StartInfo = processStartInfo;
                process.Start();
                process.WaitForExit();

                #region 用于查看是否返回错误信息
                //StreamReader srone = process.StandardError;
                //StreamReader srtwo = process.StandardOutput;
                //string ss1 = srone.ReadToEnd();
                //string ss2 = srtwo.ReadToEnd();
                //srone.Close();
                //srone.Dispose();
                //srtwo.Close();
                //srtwo.Dispose(); 
                #endregion

                process.Close();
                process.Dispose();

                flag = true;
            }
            catch
            {
                flag = false;
            }
            return flag;
        }


        /// <summary>
        /// 获取命令行参数
        /// </summary>
        /// <param name="htmlPath"></param>
        /// <param name="savePath"></param>
        /// <returns></returns>
        private string GetArguments(string htmlPath, string savePath)
        {
            if (string.IsNullOrEmpty(htmlPath))
            {
                throw new Exception("HTML local path or network address can not be empty.");
            }

            if (string.IsNullOrEmpty(savePath))
            {
                throw new Exception("The path saved by the PDF document can not be empty.");
            }

            StringBuilder stringBuilder = new StringBuilder();
            //stringBuilder.Append(" --page-height 100 ");        //页面高度100mm
            //stringBuilder.Append(" --page-width 100 ");         //页面宽度100mm
            //stringBuilder.Append(" --header-center 我是页眉 ");  //设置居中显示页眉
            //stringBuilder.Append(" --header-line ");         //页眉和内容之间显示一条直线
            //stringBuilder.Append(" --footer-center \"Page [page] of [topage]\" ");    //设置居中显示页脚
            //stringBuilder.Append(" --footer-line ");       //页脚和内容之间显示一条直线
            stringBuilder.Append(" " + htmlPath + " ");       //本地 HTML 的文件路径或网页 HTML 的URL地址
            stringBuilder.Append(" " + savePath + " ");       //生成的 PDF 文档的保存路径
            return stringBuilder.ToString();
        }


        /// <summary>
        /// 验证保存路径
        /// </summary>
        /// <param name="savePath"></param>
        private void CheckFilePath(string savePath)
        {
            string ext = string.Empty;
            string path = string.Empty;
            string fileName = string.Empty;

            ext = Path.GetExtension(savePath);
            if (string.IsNullOrEmpty(ext) || ext.ToLower() != ".pdf")
            {
                throw new Exception("Extension error:This method is used to generate PDF files.");
            }

            fileName = Path.GetFileName(savePath);
            if (string.IsNullOrEmpty(fileName))
            {
                throw new Exception("File name is empty.");
            }

            try
            {
                path = savePath.Substring(0, savePath.IndexOf(fileName));
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
            }
            catch
            {
                throw new Exception("The file path does not exist.");
            }
        }

        /// <summary>
        /// HTML文本内容转HTML文件
        /// </summary>
        /// <param name="strHtml">HTML文本内容</param>
        /// <returns>HTML文件的路径</returns>
        public string HtmlTextConvertFile(string strHtml)
        {
            if (string.IsNullOrEmpty(strHtml))
            {
                throw new Exception("HTML text content cannot be empty.");
            }

            try
            {
                string appRootPath = AppDomain.CurrentDomain.BaseDirectory;
                var fullpath = appRootPath + "/HtmlTemplate";
                if (!Directory.Exists(fullpath))
                {
                    Directory.CreateDirectory(fullpath);
                }
                string fileName = fullpath + "/" + DateTime.Now.ToString("yyyyMMddHHmmss") + new Random().Next(1000, 10000) + ".html";
                FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.Default);
                streamWriter.Write(strHtml);
                streamWriter.Flush();

                streamWriter.Close();
                streamWriter.Dispose();
                fileStream.Close();
                fileStream.Dispose();
                return fileName;
            }
            catch
            {
                throw new Exception("HTML text content error.");
            }
        }
    }
}