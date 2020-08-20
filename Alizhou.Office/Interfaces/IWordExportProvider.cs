using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Alizhou.Office.Model;

namespace Alizhou.Office.Interfaces
{
    public interface IWordExportProvider
    {
        /// <summary>
        /// 根据模板导出Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templatePath"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        AlizhouWord ExportFromTemplate<T>(string templatePath, T data) where T : IWordExportTemplate;
        /// <summary>
        /// 根据模板导出Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templatePath"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        Task<AlizhouWord> ExportFromTemplateAsync<T>(string templatePath, T data) where T : IWordExportTemplate;
    }
}
