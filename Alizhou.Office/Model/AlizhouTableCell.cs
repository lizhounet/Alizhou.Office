using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace Alizhou.Office.Model
{
    public class AlizhouTableCell
    {
        public List<AlizhouParagraph> Paragraphs { get; set; }
        /// <summary>
        /// 单元格宽
        /// </summary>
        public double Width { get; set; } = 200;
        /// <summary>
        /// 单元格填充角色
        /// </summary>
        public Color FillColor { set; get; } = Color.Empty;
        /// <summary>
        /// 单元格垂直方式(默认垂直居中)
        /// </summary>
        public VerticalAlignment VerticalAlignment { set; get; } = VerticalAlignment.Center;
    }
}
