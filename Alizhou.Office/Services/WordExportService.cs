using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Alizhou.Office.Services
{
    public class WordExportService : IWordExportService
    {
        private readonly IWordExportProvider exportProvider;
        public WordExportService(IWordExportProvider _exportProvider)
        {
            exportProvider = _exportProvider;
        }
        public AlizhouWord TemplateCreateWord<T>(string templatePath, T data) where T : IWordExportTemplate
        {
            return exportProvider.ExportFromTemplate(templatePath, data);
        }

        public IEnumerable<AlizhouWord> TemplateCreateWord<T>(string templatePath, IEnumerable<T> data) where T : IWordExportTemplate
        {
            var words = new List<AlizhouWord>();
            foreach (var item in data)
                words.Add(exportProvider.ExportFromTemplate(templatePath, item));
            return words;
        }

        public async Task<AlizhouWord> TemplateCreateWordAsync<T>(string templatePath, T data) where T : IWordExportTemplate
        {
            return await exportProvider.ExportFromTemplateAsync(templatePath, data);
        }

        public async Task<IEnumerable<AlizhouWord>> TemplateCreateWordAsync<T>(string templatePath, IEnumerable<T> data) where T : IWordExportTemplate
        {
            var words = new List<AlizhouWord>();
            foreach (var item in data)
                words.Add(await exportProvider.ExportFromTemplateAsync(templatePath, item));
            return words;
        }
    }
}
