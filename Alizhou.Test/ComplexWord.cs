using Alizhou.Office.Interfaces;
using Alizhou.Office.Provider;
using Alizhou.Office.Services;
using Alizhou.Test.TestModel;
using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace Alizhou.Test
{
    public class ComplexWord
    {
        public static void ComplexTemplateWrod()
        {
            string basePath = Environment.CurrentDirectory;
            string templateUrl = $"{basePath}/template/word/TemplateWrod.docx";

            ComplexTemplateWrod templateWrod = new ComplexTemplateWrod {
                Catalog = "1、实测实量评估结果	2\n2、质量风险评估结果  2\n3、安全文明评估结果  2\n4、管理动作评估结果  2"
            };
            IWordExportService wordExportService = new WordExportService(new WordExportProvider());
            var word = wordExportService.TemplateCreateWord(templateUrl, templateWrod);
            File.WriteAllBytes($"{basePath}/{DateTime.Now.ToString("yyyyMMddHHmmss")}ComplexTemplateWrod.docx", word.WordBytes);
        }
    }
}
