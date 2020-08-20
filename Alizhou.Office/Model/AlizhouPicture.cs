using Alizhou.Office.Interfaces;
using Novacode;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Alizhou.Office.Model
{
    public class AlizhouPicture : IWordElement
    {
        /// <summary>
        /// 图片流数据
        /// </summary>
        public Stream PictureData { get; set; }

        /// <summary>
        /// 图片绝对地址（如果PictureData不为空则不用传）
        /// </summary>
        public string PictureUrl { get; set; }
        /// <summary>
        /// 图片宽度 默认300
        /// </summary>
        public int Width { get; set; } = 300;

        /// <summary>
        /// 图片高度 默认200
        /// </summary>
        public int Height { get; set; } = 200;
    }

   
   
}
