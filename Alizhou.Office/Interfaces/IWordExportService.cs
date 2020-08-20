using Alizhou.Office.Model;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Alizhou.Office.Interfaces
{
    public interface IWordExportService
    {
        /// <summary>
        /// 根据模板生成Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templatePath"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        Task<AlizhouWord> TemplateCreateWordAsync<T>(string templatePath, T data) where T : IWordExportTemplate;
        /// <summary>
        /// 根据模板生成Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templatePath"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        AlizhouWord TemplateCreateWord<T>(string templatePath, T data) where T : IWordExportTemplate;

        /// <summary>
        /// 根据模板生成Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templatePath"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        Task<IEnumerable<AlizhouWord>> TemplateCreateWordAsync<T>(string templatePath, IEnumerable<T> data) where T : IWordExportTemplate;
        /// <summary>
        /// 根据模板生成Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templatePath"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        IEnumerable<AlizhouWord> TemplateCreateWord<T>(string templatePath, IEnumerable<T> data) where T : IWordExportTemplate;
    }
}
