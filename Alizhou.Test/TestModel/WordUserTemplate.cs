using Alizhou.Office.Attribute;
using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using Novacode;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Test.TestModel
{
    public class WordUserTemplate : IWordExportTemplate
    {
        /// <summary>
        ///  默认占位符为{PropertyName}
        /// </summary>


        public string UserName { get; set; }
        /// <summary>
        /// 默认
        /// </summary>
        [Placeholder("{电话}")]
        public string Phone { set; get; }

        /// <summary>
        /// 表格
        /// </summary>
        public AlizhouTable Table { get; set; }
        /// <summary>
        /// 图片
        /// </summary>
        public IEnumerable<AlizhouPicture> Pictures { get; set; }
    }
}
