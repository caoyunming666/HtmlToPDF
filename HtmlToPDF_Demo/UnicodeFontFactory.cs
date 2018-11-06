using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace HtmlToPDF_Demo
{
    public class UnicodeFontFactory : FontFactoryImp
    {
        private static readonly string arialFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                  "arialuni.ttf");//arial unicode MS是完整的unicode字型。
        private static readonly string biaoKaiTiPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
          "KAIU.TTF");//标楷体

        //    private static readonly string arialFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
        //"SIMYOU.TTF");//arial unicode MS是完整的unicode字型。
        //https://blog.csdn.net/gf771115/article/details/50556761 字体设置参考地址
        //BaseFont baseFont = BaseFont.createFont("C:/Windows/Fonts/SIMYOU.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);    

        public override Font GetFont(string fontname, string encoding, bool embedded, float size, int style, BaseColor color,
            bool cached)
        {
            BaseFont baseFont = BaseFont.CreateFont(biaoKaiTiPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            return new Font(baseFont, size, style, color);
        }
    }
}