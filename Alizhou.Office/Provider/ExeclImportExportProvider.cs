using Alizhou.Office.Helper;
using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Alizhou.Office.Provider
{
    public class ExeclImportExportProvider : IExeclImportExportProvider
    {
        public AlizhouExecl Export<T>(ICollection<T> data)
        {
            return new AlizhouExecl()
            {
                WordBytes = EPPlusHelper.Export(data)
            };
        }

        public async Task<AlizhouExecl> ExportAsync<T>(ICollection<T> data)
         => await Task.Run(() => Export(data));

        public ICollection<ImportT> Import<ImportT>(Stream stream) where ImportT : new()
        {
            return EPPlusHelper.Import<ImportT>(stream);
        }

        public async Task<ICollection<ImportT>> ImportAsync<ImportT>(Stream stream) where ImportT : new()
         => await Task.Run(() => Import<ImportT>(stream));
    }
}
