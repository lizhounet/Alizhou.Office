using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using Alizhou.Office.Provider;
using Alizhou.Office.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace Alizhou.Test
{
    public class EvaluationReport : IWordExportTemplate
    {
        /// <summary>
        /// 项目名称
        /// </summary>
        public string ProjectName { set; get; }
        /// <summary>
        /// 项目地址
        /// </summary>
        public string ProjectAdress { set; get; }
        /// <summary>
        /// 城市
        /// </summary>
        public string City { set; get; }
        /// <summary>
        /// 报告时间
        /// </summary>
        public string ReportDate { set; get; }
        /// <summary>
        /// 评估类型
        /// </summary>
        public string EvaluationType { set; get; }
        /// <summary>
        /// 评估标段
        /// </summary>

        public string BidSection { set; get; }
        /// <summary>
        /// 组长
        /// </summary>
        public string GroupLeader { set; get; }
        /// <summary>
        /// 组员
        /// </summary>
        public string GroupMember { set; get; }
        /// <summary>
        /// 监理单位
        /// </summary>
        public string SupervisionUnit { set; get; }
        /// <summary>
        /// 施工单位
        /// </summary>
        public string ConstructionUnit { set; get; }
        /// <summary>
        /// 项目负责人
        /// </summary>
        public string ProjectPersonInchargeName { set; get; }
        /// <summary>
        /// 项目组成
        /// </summary>
        public string Composition { set; get; }
        /// <summary>
        /// 测区说明
        /// </summary>

        public string MeasuringareaDescription { set; get; }
        /// <summary>
        /// 综合评估结果
        /// </summary>

        public AlizhouTable ComplexResult { set; get; }
        /// <summary>
        /// 分项评估结果
        /// </summary>
        public AlizhouComplex SubOptionResult { set; get; }
    }
    public class EvaluationReportUniversal
    {
        public static void EvaluationReportTemplateWrod()
        {
            string basePath = Environment.CurrentDirectory;
            string templateUrl = $"{basePath}/template/word/EvaluationReportUniversal.docx";
            var templateWrod = new EvaluationReport();
            templateWrod.City = "成都";
            templateWrod.BidSection = "一标段";
            templateWrod.ReportDate = DateTime.Now.ToString("yyyy-MM-dd");
            templateWrod.ProjectName = "蓝光空港国际城住宅项目";
            templateWrod.ProjectAdress = "成都郫县";
            templateWrod.EvaluationType = "专项评估";
            templateWrod.GroupLeader = "张三";
            templateWrod.GroupMember = "李四；王五";
            templateWrod.SupervisionUnit = "成新设计院咨询部";
            templateWrod.ConstructionUnit = "蓝光集团";
            templateWrod.ProjectPersonInchargeName = "刘晨";
            templateWrod.Composition = "项目组成";
            templateWrod.MeasuringareaDescription = "精装-3标段-5栋6单元1号\n精装-3标段-6栋6单元1号\n精装-3标段-7栋6单元24号\n精装-3标段-5栋6单元12号";
            #region 综合评估结果
            //综合评估结果
            var complexResult = new Dictionary<string, string>() { };
            complexResult.Add("实测实量(xx %)", "95%");
            complexResult.Add("质量风险(xx %)", "65%");
            complexResult.Add("安全文明(xx %)", "75%");
            complexResult.Add("管理行为(xx %)", "85%");
            var complexResultTable = new AlizhouTable(2 + complexResult.Count, 4);
            //处理标题
            complexResultTable.Rows[0].Height = 40;
            complexResultTable.Rows[0].Cells[0].Paragraphs[0].Run.Text = "项目标段名称";
            complexResultTable.Rows[0].Cells[0].Paragraphs[0].Run.IsBold = true;
            complexResultTable.Rows[0].Cells[0].FillColor = Color.FromArgb(242, 242, 242);


            complexResultTable.Rows[0].Cells[1].Paragraphs[0].Run.Text = "分项名称";
            complexResultTable.Rows[0].Cells[1].Paragraphs[0].Run.IsBold = true;
            complexResultTable.Rows[0].Cells[1].FillColor = Color.FromArgb(242, 242, 242);

            complexResultTable.Rows[0].Cells[2].Paragraphs[0].Run.Text = "分项评估结果";
            complexResultTable.Rows[0].Cells[2].Paragraphs[0].Run.IsBold = true;
            complexResultTable.Rows[0].Cells[2].FillColor = Color.FromArgb(242, 242, 242);

            complexResultTable.Rows[0].Cells[3].Paragraphs[0].Run.Text = "综合评估结果";
            complexResultTable.Rows[0].Cells[3].Paragraphs[0].Run.IsBold = true;
            complexResultTable.Rows[0].Cells[3].FillColor = Color.FromArgb(242, 242, 242);
            {
                int index = 1;
                foreach (var item in complexResult)
                {

                    //填充数据
                    complexResultTable.Rows[index].Height = 37;
                    complexResultTable.Rows[index].Cells[0].Paragraphs[0].Run.Text = "一标段";
                    complexResultTable.Rows[index].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
                    complexResultTable.Rows[index].Cells[1].Paragraphs[0].Run.Text = item.Key;
                    complexResultTable.Rows[index].Cells[2].Paragraphs[0].Run.Pictures.Add(new AlizhouPicture { PictureUrl = "D://图片1.png" });

                    complexResultTable.Rows[index].Cells[3].Paragraphs[0].Run.Text = "99%";
                    index++;
                }
            }
            complexResultTable.Rows[complexResultTable.RowCount - 1].Height = 40;
            complexResultTable.Rows[complexResultTable.RowCount - 1].Cells[0].Paragraphs[0].Run.Text = "备注";
            complexResultTable.Rows[complexResultTable.RowCount - 1].Cells[0].Paragraphs[0].Run.IsBold = true;
            complexResultTable.Rows[complexResultTable.RowCount - 1].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            complexResultTable.Rows[complexResultTable.RowCount - 1].Cells[1].Paragraphs[0].Run.Text = "综合评估结果=各维度成绩*各维权重";


            complexResultTable.MergeCellsInColumn(0, 1, complexResultTable.RowCount - 2);
            complexResultTable.MergeCellsInColumn(3, 1, complexResultTable.RowCount - 1);
            complexResultTable.MergeCellsInRow(complexResultTable.RowCount - 1, 1, 3);
            templateWrod.ComplexResult = complexResultTable;
            #endregion
            #region 分项评估结果
            templateWrod.SubOptionResult = new AlizhouComplex();
            templateWrod.SubOptionResult.Elements.Add(new AlizhouParagraph { Alignment = Novacode.Alignment.left, Run = new AlizhouRun { IsBold = true, Text = "1、实测实量评估结果", FontFamily = "宋体", FontSize = 12 } });
            templateWrod.SubOptionResult.Elements.Add(new AlizhouParagraph { Run = new AlizhouRun { Text = "表 1  实测实量评估结果分析表", FontFamily = "黑体", FontSize = 8 } });
            //实量实测表格
            //var measuredTable = new AlizhouTable(fbgc.Length + 5, 8);
            //measuredTable.Rows[0].Cells[0].Paragraphs[0].Run.Text = templateWrod.ProjectName;
            //measuredTable.Rows[0].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[0].Cells[0].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[0].Cells[1].Paragraphs[0].Run.Text = "分部工程";
            //measuredTable.Rows[0].Cells[1].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[0].Cells[1].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[0].Cells[2].Paragraphs[0].Run.Text = "较好指标";
            //measuredTable.Rows[0].Cells[2].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[0].Cells[2].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[0].Cells[4].Paragraphs[0].Run.Text = "一般指标";
            //measuredTable.Rows[0].Cells[4].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[0].Cells[4].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[0].Cells[6].Paragraphs[0].Run.Text = "较差指标";
            //measuredTable.Rows[0].Cells[6].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[0].Cells[6].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[1].Height = 60;
            //measuredTable.Rows[1].Cells[2].Paragraphs[0].Run.Text = "名称";
            //measuredTable.Rows[1].Cells[2].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[1].Cells[2].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[1].Cells[3].Paragraphs[0].Run.Text = "合格率≥90%";
            //measuredTable.Rows[1].Cells[3].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[1].Cells[3].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[1].Cells[4].Paragraphs[0].Run.Text = "名称";
            //measuredTable.Rows[1].Cells[4].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[1].Cells[4].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[1].Cells[5].Paragraphs[0].Run.Text = "90%＞合格率＞70%";
            //measuredTable.Rows[1].Cells[5].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[1].Cells[5].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[1].Cells[6].Paragraphs[0].Run.Text = "名称";
            //measuredTable.Rows[1].Cells[6].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[1].Cells[6].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[1].Cells[7].Paragraphs[0].Run.Text = "合格率≤70%";
            //measuredTable.Rows[1].Cells[7].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[1].Cells[7].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[measuredTable.RowCount - 3].Cells[1].Paragraphs[0].Run.Text = "实测实量评估结果";
            //measuredTable.Rows[measuredTable.RowCount - 3].Cells[1].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[measuredTable.RowCount - 3].Cells[1].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[measuredTable.RowCount - 3].Cells[5].Paragraphs[0].Run.Text = "测量点总数";
            //measuredTable.Rows[measuredTable.RowCount - 3].Cells[5].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[measuredTable.RowCount - 3].Cells[5].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[measuredTable.RowCount - 2].Cells[5].Paragraphs[0].Run.Text = "合格点总数";
            //measuredTable.Rows[measuredTable.RowCount - 2].Cells[5].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[measuredTable.RowCount - 2].Cells[5].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[measuredTable.RowCount - 1].Height = 50;
            //measuredTable.Rows[measuredTable.RowCount - 1].Cells[0].Paragraphs[0].Run.Text = "备注";
            //measuredTable.Rows[measuredTable.RowCount - 1].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            //measuredTable.Rows[measuredTable.RowCount - 1].Cells[0].Paragraphs[0].Run.IsBold = true;

            //measuredTable.Rows[measuredTable.RowCount - 1].Cells[3].Paragraphs[0].Run.Text = "实测实量评估结果=实测合格点总点数/实测点总点数*100%";
            //{
            //    int index = 2;
            //    foreach (var item in fbgc)
            //    {
            //        measuredTable.Rows[index].Cells[1].Paragraphs[0].Run.Text = item;
            //        measuredTable.Rows[index].Cells[1].Paragraphs[0].Run.IsBold = true;
            //        index++;
            //    }
            //}

            //measuredTable.MergeCellsInRow(0, 2, 3);
            //measuredTable.MergeCellsInRow(0, 3, 4);
            //measuredTable.MergeCellsInRow(0, 4, 5);
            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 3, 1, 2);
            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 3, 2, 3);
            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 3, 3, 4);
            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 2, 1, 2);
            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 2, 2, 3);
            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 2, 3, 4);

            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 1, 0, 2);
            //measuredTable.MergeCellsInRow(measuredTable.RowCount - 1, 1, 5);

            //measuredTable.MergeCellsInColumn(0, 0, measuredTable.RowCount - 2);
            //measuredTable.MergeCellsInColumn(1, 0, 1);

            //measuredTable.MergeCellsInColumn(3, measuredTable.RowCount - 3, measuredTable.RowCount - 2);
            //measuredTable.MergeCellsInColumn(1, measuredTable.RowCount - 3, measuredTable.RowCount - 2);


            var measuredTable = new AlizhouTable(3 + 5, 5);
            measuredTable.Rows[0].Height = 40;
            measuredTable.Rows[0].Cells[0].Paragraphs[0].Run.Text = templateWrod.ProjectName;
            measuredTable.Rows[0].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[0].Cells[0].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[0].Cells[1].Paragraphs[0].Run.Text = "分部工程";
            measuredTable.Rows[0].Cells[1].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[0].Cells[1].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[0].Cells[2].Paragraphs[0].Run.Text = "名称";
            measuredTable.Rows[0].Cells[2].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[0].Cells[2].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[0].Cells[3].Paragraphs[0].Run.Text = "合格率";
            measuredTable.Rows[0].Cells[3].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[0].Cells[3].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[0].Cells[4].Paragraphs[0].Run.Text = "指标";
            measuredTable.Rows[0].Cells[4].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[0].Cells[4].Paragraphs[0].Run.IsBold = true;
            {
                int index = 1;
                foreach (var item in new string[] { "抹灰工程", "设备安装", "门窗工程" })
                {
                    measuredTable.Rows[index].Height = 40;
                    measuredTable.Rows[index].Cells[1].Paragraphs[0].Run.Text = item;
                    measuredTable.Rows[index].Cells[2].Paragraphs[0].Run.Text = item + index;
                    index++;
                }
            }


            measuredTable.Rows[measuredTable.RowCount - 1].Cells[0].Paragraphs[0].Run.Text = "指标说明";
            measuredTable.Rows[measuredTable.RowCount - 1].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[measuredTable.RowCount - 1].Cells[0].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[measuredTable.RowCount - 1].Cells[2].Paragraphs[0].Run.Text = "较好-合格率≥90%\n 一般-90%＞合格率＞70%\n较差-合格率≤70%";

            measuredTable.Rows[measuredTable.RowCount - 2].Cells[0].Paragraphs[0].Run.Text = "备注";
            measuredTable.Rows[measuredTable.RowCount - 2].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[measuredTable.RowCount - 2].Cells[0].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[measuredTable.RowCount - 2].Cells[2].Paragraphs[0].Run.Text = "实测实量评估结果=实测合格点总点数/实测点总点数*100%";

            measuredTable.Rows[measuredTable.RowCount - 4].Cells[0].Paragraphs[0].Run.Text = "实测实量评估结果";
            measuredTable.Rows[measuredTable.RowCount - 4].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[measuredTable.RowCount - 4].Cells[0].Paragraphs[0].Run.IsBold = true;


            measuredTable.Rows[measuredTable.RowCount - 4].Cells[1].Paragraphs[0].Run.Text = "99.9%";

            measuredTable.Rows[measuredTable.RowCount - 4].Cells[3].Paragraphs[0].Run.Text = "测量点总数";
            measuredTable.Rows[measuredTable.RowCount - 4].Cells[3].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[measuredTable.RowCount - 4].Cells[3].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[measuredTable.RowCount - 4].Cells[4].Paragraphs[0].Run.Text = "200";

            measuredTable.Rows[measuredTable.RowCount - 3].Cells[3].Paragraphs[0].Run.Text = "合格点总数";
            measuredTable.Rows[measuredTable.RowCount - 3].Cells[3].FillColor = Color.FromArgb(242, 242, 242);
            measuredTable.Rows[measuredTable.RowCount - 3].Cells[3].Paragraphs[0].Run.IsBold = true;

            measuredTable.Rows[measuredTable.RowCount - 3].Cells[4].Paragraphs[0].Run.Text = "180";

            measuredTable.MergeCellsInColumn(0, 0, measuredTable.RowCount - 5);

            measuredTable.MergeCellsInRow(measuredTable.RowCount - 1, 0, 1);
            measuredTable.MergeCellsInRow(measuredTable.RowCount - 1, 1, 3);

            measuredTable.MergeCellsInRow(measuredTable.RowCount - 2, 0, 1);
            measuredTable.MergeCellsInRow(measuredTable.RowCount - 2, 1, 3);

            measuredTable.MergeCellsInRow(measuredTable.RowCount - 3, 1, 2);
            measuredTable.MergeCellsInRow(measuredTable.RowCount - 4, 1, 2);

            measuredTable.MergeCellsInColumn(0, measuredTable.RowCount - 4, measuredTable.RowCount - 3);
            measuredTable.MergeCellsInColumn(1, measuredTable.RowCount - 4, measuredTable.RowCount - 3);
            measuredTable.MergeCellsInColumn(2, measuredTable.RowCount - 4, measuredTable.RowCount - 3);

            templateWrod.SubOptionResult.Elements.Add(measuredTable);

            //风险评估部分
            templateWrod.SubOptionResult.Elements.Add(new AlizhouParagraph { Alignment = Novacode.Alignment.left, Run = new AlizhouRun { IsBold = true, Text = "2、质量风险评估结果", FontFamily = "宋体", FontSize = 12 } });
            templateWrod.SubOptionResult.Elements.Add(new AlizhouParagraph { Run = new AlizhouRun { Text = "表 1  质量风险评估结果分析表", FontFamily = "黑体", FontSize = 8 } });
            var riskTable = new AlizhouTable(3 + 2, 3);
            riskTable.Rows[0].Height = 40;
            riskTable.Rows[0].Cells[0].Paragraphs[0].Run.Text = "质量风险评分汇总";
            riskTable.Rows[0].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            riskTable.Rows[0].Cells[0].Paragraphs[0].Run.IsBold = true;
            riskTable.Rows[1].Height = 40;
            riskTable.Rows[1].Cells[0].Paragraphs[0].Run.Text = "分项工程";
            riskTable.Rows[1].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            riskTable.Rows[1].Cells[0].Paragraphs[0].Run.IsBold = true;
            riskTable.Rows[1].Cells[1].Paragraphs[0].Run.Text = "分项合格率";
            riskTable.Rows[1].Cells[1].FillColor = Color.FromArgb(242, 242, 242);
            riskTable.Rows[1].Cells[1].Paragraphs[0].Run.IsBold = true;
            riskTable.Rows[1].Cells[2].Paragraphs[0].Run.Text = "质量风险评估结果";
            riskTable.Rows[1].Cells[2].FillColor = Color.FromArgb(242, 242, 242);
            riskTable.Rows[1].Cells[2].Paragraphs[0].Run.IsBold = true;
            {
                int index = 2;
                foreach (var item in new string[] { "渗漏", "空鼓/开裂" })
                {
                    riskTable.Rows[index].Height = 40;
                    riskTable.Rows[index].Cells[0].Paragraphs[0].Run.Text = item;
                    riskTable.Rows[index].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
                    riskTable.Rows[index].Cells[0].Paragraphs[0].Run.IsBold = true;

                    riskTable.Rows[index].Cells[1].Paragraphs[0].Run.Text = "12%";
                    riskTable.Rows[index].Cells[2].Paragraphs[0].Run.Text = "85%";
                    index++;
                }
            }

            riskTable.Rows[riskTable.RowCount - 1].Cells[0].Paragraphs[0].Run.Text = "备注";
            riskTable.Rows[riskTable.RowCount - 1].Cells[0].FillColor = Color.FromArgb(242, 242, 242);
            riskTable.Rows[riskTable.RowCount - 1].Cells[0].Paragraphs[0].Run.IsBold = true;

            riskTable.Rows[riskTable.RowCount - 1].Cells[1].Paragraphs[0].Run.Text = "质量风险评估结果=实得分/应得分*100%。";

            riskTable.MergeCellsInRow(0, 0, 2);
            riskTable.MergeCellsInColumn(2, 2, riskTable.RowCount - 2);
            riskTable.MergeCellsInRow(riskTable.RowCount - 1, 1, 2);

            templateWrod.SubOptionResult.Elements.Add(riskTable);
            #endregion

            IWordExportService wordExportService = new WordExportService(new WordExportProvider());
            var word = wordExportService.TemplateCreateWord(templateUrl, templateWrod);
            File.WriteAllBytes($@"{basePath}..\..\..\..\OutPut\word\OutEvaluationReportUniversal.docx", word.WordBytes);
        }
    }
}
