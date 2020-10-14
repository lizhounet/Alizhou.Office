using Alizhou.Office.Interfaces;
using Alizhou.Office.Provider;
using Alizhou.Office.Services;
using EPPlus.Core.Extensions.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Alizhou.Test.Execl
{
    /// <summary>
    /// execl导入导出
    /// </summary>
    public class ExeclImportExportTest
    {
        public static void Test()
        {
            string basePath = Environment.CurrentDirectory;
            var list = new List<PersonDto>();
            for (int i = 0; i < 100; i++)
            {
                list.Add(new PersonDto
                {
                    Name = $"张三{i}",
                    Age = 18,
                    Adress = $"地址{i}",
                    Phone = $"17783042962"
                });
            }
            IExeclImportExportService execlImportExport = new ExeclImportExportService(new ExeclImportExportProvider());
            var alizhouExecl = execlImportExport.Export(list);
            string path = $@"{basePath}..\..\..\..\OutPut\execl\ExeclImportExportTest.xlsx";
            File.WriteAllBytes(path, alizhouExecl.WordBytes);
            var data = execlImportExport.Import<PersonDto>(File.OpenRead(path));
            foreach (var item in data)
            {
                Console.WriteLine(item.ToString());
            }
        }
    }
    public class PersonDto
    {
        [ExcelTableColumn("姓名")]
        [Required(ErrorMessage = "姓名必填")]

        public string Name { get; set; }
        [ExcelTableColumn("年龄")]
        [Range(1, 25, ErrorMessage = "年龄超过25岁的MM不要")]//范围判断
        public int Age { get; set; }

        [ExcelTableColumn(3)]//根据索引取
        public string Adress { get; set; }
        [ExcelTableColumn("电话")]
        [MaxLength(11, ErrorMessage = "手机号 长度不能超过11")]
        public string Phone { set; get; }
        public override string ToString()
        {
            return $"姓名：{Name}；年龄：{Age}；地址：{Adress}；电话：{Phone}";
        }
    }
}
