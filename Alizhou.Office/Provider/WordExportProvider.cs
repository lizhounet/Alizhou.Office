using Alizhou.Office.Attribute;
using Alizhou.Office.Enum;
using Alizhou.Office.Extensions;
using Alizhou.Office.Helper;
using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using Novacode;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Alizhou.Office.Provider
{
    public class WordExportProvider : IWordExportProvider
    {
        public AlizhouWord ExportFromTemplate<T>(string templatePath, T data) where T : IWordExportTemplate
        {
            var word = DocXHelper.GetDocX(templatePath);
            ReplacePlaceholders(word, data);
            return new AlizhouWord()
            {
                WordBytes = word.ToBytes()
            };


        }
        public async Task<AlizhouWord> ExportFromTemplateAsync<T>(string templatePath, T data) where T : IWordExportTemplate
        {
            return await Task.Run(() =>
            {
                var word = DocXHelper.GetDocX(templatePath);
                ReplacePlaceholders(word, data);
                return new AlizhouWord()
                {
                    WordBytes = word.ToBytes()
                };
            });
        }
        /// <summary>
        /// 替换占位符
        /// </summary>
        /// <param name="word"></param>
        private void ReplacePlaceholders<T>(DocX word, T wordData)
            where T : IWordExportTemplate
        {
            if (word == null)
                throw new ArgumentNullException("word");
            var placeholders = wordData.GetReplacements();
            if (placeholders == null) throw new Exception("实体中没有可替换的属性");
            DocXHelper.ReplacePlaceholdersInWord(word, placeholders);
        }
    }
}
