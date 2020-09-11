using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
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
        public AlizhouComplex SLFX { set; get; } = new AlizhouComplex();
    }
}
