using Alizhou.Office.Interfaces;
using Alizhou.Office.Provider;
using Alizhou.Office.Services;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Extensions
{
    /// <summary>
    /// .netcore 依赖注入使用
    /// </summary>
    public static class OfficeServiceCollectionExtensions
    {
        public static void AddAlizhouOffice(this IServiceCollection services)
        {
            services.AddTransient<IWordExportProvider, WordExportProvider>();
            services.AddTransient<IWordExportService, WordExportService>();
        }
    }

}
