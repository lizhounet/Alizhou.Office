using Alizhou.Office.Enum;
using Alizhou.Office.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Model
{
    public class PlaceholderEntity
    {
        /// <summary>
        /// 占位符
        /// </summary>
        public string Placeholder { set; get; }
        /// <summary>
        /// 数据类型
        /// </summary>
        public PlaceholderType PlaceholderType { set; get; }
        /// <summary>
        /// 数据
        /// </summary>
        public IWordElement Data { set; get; }
        /// <summary>
        /// 图片
        /// </summary>
        public IEnumerable<AlizhouPicture> Pictures { set; get; }
    }
}
