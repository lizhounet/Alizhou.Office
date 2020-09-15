using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Alizhou.Office.Services
{
    public class ExeclImportExportService : IExeclImportExportService
    {
        private readonly IExeclImportExportProvider execlImportExport;
        public ExeclImportExportService(IExeclImportExportProvider _execlImportExport)
        {
            execlImportExport = _execlImportExport;
        }
        public AlizhouExecl Export<ExportT>(ICollection<ExportT> data)
        {
            return execlImportExport.Export(data);
        }

        public async Task<AlizhouExecl> ExportAsync<ExportT>(ICollection<ExportT> data)
        {
            return await execlImportExport.ExportAsync(data);
        }

        public ICollection<ImportT> Import<ImportT>(Stream stream) where ImportT : new()
        {
            return execlImportExport.Import<ImportT>(stream);
        }

        public async Task<ICollection<ImportT>> ImportAsync<ImportT>(Stream stream) where ImportT : new()
        {
            return await execlImportExport.ImportAsync<ImportT>(stream);
        }
    }
}
