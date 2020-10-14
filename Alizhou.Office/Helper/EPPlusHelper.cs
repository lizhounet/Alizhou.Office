using EPPlus.Core.Extensions;
using EPPlus.Core.Extensions.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Text;

namespace Alizhou.Office.Helper
{
    internal class EPPlusHelper
    {

        public static byte[] Export<T>(ICollection<T> data)
        {
            if (data == null) throw new ArgumentNullException("data不能为空");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = data.ToExcelPackage();
            var worksheet = excelPackage.GetWorksheet(0);
            worksheet.Cells.AutoFitColumns();
            using (MemoryStream ms = new MemoryStream())
            {
                excelPackage.SaveAs(ms);
                return ms.ToArray();
            }
            //   return data.ToXlsx();
        }

        public static ICollection<T> Import<T>(Stream stream) where T : new()
        {
            if (stream == null) throw new ArgumentNullException("stream不能为空");
            ExcelPackage package = new ExcelPackage(stream);
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];
            return sheet.ToList<T>();
        }
    }
}
