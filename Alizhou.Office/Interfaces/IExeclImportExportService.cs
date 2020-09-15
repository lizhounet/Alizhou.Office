using Alizhou.Office.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Alizhou.Office.Interfaces
{
    public interface IExeclImportExportService
    {
        /// <summary>
        /// 根据数据导出execl
        /// </summary>
        /// <typeparam name="T">导出execl实体</typeparam>
        /// <param name="data">源数据</param>
        /// <returns></returns>
        AlizhouExecl Export<ExportT>(ICollection<ExportT> data);
        /// <summary>
        /// 根据数据导出execl
        /// </summary>
        /// <typeparam name="T">导出execl实体</typeparam>
        /// <param name="data">源数据</param>
        /// <returns></returns>
        Task<AlizhouExecl> ExportAsync<ExportT>(ICollection<ExportT> data);
        /// <summary>
        /// 导出execl
        /// </summary>
        /// <typeparam name="ImportT">导出模板实体</typeparam>
        /// <param name="stream"></param>
        /// <returns></returns>
        ICollection<ImportT> Import<ImportT>(Stream stream) where ImportT : new();
        /// <summary>
        /// 导出execl
        /// </summary>
        /// <typeparam name="ImportT">导出模板实体</typeparam>
        /// <param name="stream"></param>
        /// <returns></returns>
        Task<ICollection<ImportT>> ImportAsync<ImportT>(Stream stream) where ImportT : new();
    }
}
