using Alizhou.Office.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Test.TestModel
{
    public class ComplexTemplateWrod: IWordExportTemplate
    {
        /// <summary>
        /// 目录
        /// </summary>
        public string Catalog { set; get; }
    }
}
