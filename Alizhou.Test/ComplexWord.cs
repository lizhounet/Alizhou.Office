using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using Alizhou.Office.Provider;
using Alizhou.Office.Services;
using Alizhou.Test.TestModel;
using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace Alizhou.Test
{
    public class ComplexWord
    {
        public static void ComplexTemplateWrod()
        {
            string basePath = Environment.CurrentDirectory;
            string templateUrl = $"{basePath}/template/word/报告模板（通用）.docx";


            string[] columnNames = { "渗漏", "空鼓/开裂", "观感质量", "成品保护", "结构安全", "违规、强条", "备注" };

            ComplexTemplateWrod templateWrod = new ComplexTemplateWrod
            {
                Catalog = "1、实测实量评估结果	2\n2、质量风险评估结果  2\n3、安全文明评估结果  2\n4、管理动作评估结果  2"
            };
            var table = new AlizhouTable(9, 3);
            //添加标题
            table.Rows[0].Cells[0].Paragraphs[0].Run.Text = "质量风险评分汇总";
            table.Rows[0].Cells[0].Paragraphs[0].Run.IsBold = true;

            table.Rows[1].Cells[0].Paragraphs[0].Run.Text = "分项工程";
            table.Rows[1].Cells[0].Paragraphs[0].Run.IsBold = true;

            table.Rows[1].Cells[1].Paragraphs[0].Run.Text = "分项合格率";
            table.Rows[1].Cells[1].Paragraphs[0].Run.IsBold = true;

            table.Rows[1].Cells[2].Paragraphs[0].Run.Text = "质量风险评估结果";
            table.Rows[1].Cells[2].Paragraphs[0].Run.IsBold = true;

            for (int i = 2; i < table.RowCount; i++)
            {
                table.Rows[i].Cells[0].Paragraphs[0].Run.Text = columnNames[i - 2];
                table.Rows[i].Cells[0].Paragraphs[0].Run.IsBold = true;

            }
            //合并单元格
            table.MergeCellsInColumn(2, 2, 7);
            table.MergeCellsInRow(0, 0, 2);

            templateWrod.SLFX.Elements.Add(new AlizhouParagraph { Run = new AlizhouRun { Text = "1、实测实量评估结果", FontSize = 20, IsBold = true}, Alignment = Alignment.left });
            templateWrod.SLFX.Elements.Add(table);
            templateWrod.SLFX.Elements.Add(new AlizhouParagraph { Run = new AlizhouRun { Text = "2、风险评估评估结果", FontSize = 20, IsBold = true }, Alignment = Alignment.left });
            templateWrod.SLFX.Elements.Add(table);
            IWordExportService wordExportService = new WordExportService(new WordExportProvider());
            var word = wordExportService.TemplateCreateWord(templateUrl, templateWrod);
            File.WriteAllBytes($"{basePath}/{DateTime.Now.ToString("yyyyMMddHHmmss")}ComplexTemplateWrod.docx", word.WordBytes);
        }
    }
}
