using Novacode;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Model
{
    /// <summary>
    /// 段落
    /// </summary>
   public class AlizhouParagraph
    {
        /// <summary>
        /// 文本
        /// </summary>
        public AlizhouRun Run { get; set; }
        /// <summary>
        /// 默认靠左
        /// </summary>
        public Alignment Alignment { get; set; } = Alignment.left;
    }
}
