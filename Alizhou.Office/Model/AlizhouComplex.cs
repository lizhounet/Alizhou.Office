using Alizhou.Office.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Model
{
    /// <summary>
    /// 复杂元素
    /// </summary>
    public class AlizhouComplex: IWordElement
    {
        /// <summary>
        /// 元素(支持AlizhouParagraph，AlizhouTable，AlizhouPicture)
        /// </summary>
        public List<IWordElement> Elements { set; get; } = new List<IWordElement>();
    }
}
