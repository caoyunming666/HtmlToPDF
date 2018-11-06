using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace HtmlToPDF_Demo
{
    /// <summary>
    /// 自定义 动作方法返回结果
    ///     使用response输出一个pdf文件
    /// </summary>
    public class BinaryContentResult : ActionResult
    {
        private readonly string contentType;
        private readonly byte[] contentBytes;

        public BinaryContentResult(byte[] contentBytes, string contentType)
        {
            this.contentBytes = contentBytes;
            this.contentType = contentType;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            var response = context.HttpContext.Response;
            response.Clear();
            response.Cache.SetCacheability(HttpCacheability.Public);
            response.ContentType = this.contentType;
            string tradeTime = DateTime.Now.ToString("yyyyMMddHHmmss", DateTimeFormatInfo.InvariantInfo);
            //下面这段加上就是一个下载页面
            response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(tradeTime + ".pdf", Encoding.UTF8));
            using (MemoryStream stream = new MemoryStream(this.contentBytes))
            {
                stream.WriteTo(response.OutputStream);
                stream.Flush();
            }
        }
    }
}