﻿using Alizhou.Office.Interfaces;
using Novacode;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Model
{
    /// <summary>
    /// 段落
    /// </summary>
   public class AlizhouParagraph: IWordElement
    {
        /// <summary>
        /// 文本
        /// </summary>
        public AlizhouRun Run { get; set; }
        /// <summary>
        /// 默认居中
        /// </summary>
        public Alignment Alignment { get; set; } = Alignment.center;
    }
}
