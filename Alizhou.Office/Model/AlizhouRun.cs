using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace Alizhou.Office.Model
{
    public class AlizhouRun
    {
        public string Text { get; set; } = "";
        public Color Color { get; set; } = Color.Black;
        public int FontSize { get; set; } = 12;
        public string FontFamily { get; set; } = "等线 (中文正文)";
        public bool IsBold { get; set; } = false;
        public List<AlizhouPicture> Pictures { get; set; } = new List<AlizhouPicture>();
    }
}
