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

namespace Alizhou.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //string basePath = Environment.CurrentDirectory;
            //string templateUrl = $"{basePath}/template/word/TemplateWrod.docx";
            //IWordExportService wordExportService = new WordExportService(new WordExportProvider());
            //WordUserTemplate userTemplate = new WordUserTemplate
            //{
            //    UserName = "小周黎",
            //    Phone = "175626565656",
            //    Pictures = new List<AlizhouPicture>() {
            //    new Office.Model.AlizhouPicture
            //    {
            //        PictureUrl = "D://11566d53-14f2-4d2a-b2e1-c47b8d319eb9.png",
            //        Width=540,
            //        Height=405,
            //    },
            //     new Office.Model.AlizhouPicture
            //    {
            //        PictureUrl = "D://图片2.png",
            //        Width=540,
            //        Height=405,
            //    }
            //    },
            //    Table = new Office.Model.AlizhouTable(3, 2)
            //    {
            //        Rows = new System.Collections.Generic.List<Office.Model.AlizhouTableRow>() {
            //            new Office.Model.AlizhouTableRow{
            //                Height=200,
            //                Cells=new System.Collections.Generic.List<Office.Model.AlizhouTableCell>(){
            //                    new Office.Model.AlizhouTableCell{
            //                        Width=100,
            //                        Paragraphs=new System.Collections.Generic.List<Office.Model.AlizhouParagraph>(){
            //                            new Office.Model.AlizhouParagraph{
            //                                Run=new Office.Model.AlizhouRun{
            //                                    Text="姓名",
            //                                     IsBold=true
            //                                }
            //                            }
            //                        }

            //                    },
            //                    new Office.Model.AlizhouTableCell{
            //                        Width=100,
            //                        Paragraphs=new System.Collections.Generic.List<Office.Model.AlizhouParagraph>(){
            //                            new Office.Model.AlizhouParagraph{
            //                                Run=new Office.Model.AlizhouRun{
            //                                    Text="年龄"
            //                                }
            //                            }
            //                        }

            //                    }
            //                },
            //            },
            //             new Office.Model.AlizhouTableRow{
            //                Height=200,
            //                Cells=new System.Collections.Generic.List<Office.Model.AlizhouTableCell>(){
            //                    new Office.Model.AlizhouTableCell{
            //                        Width=100,
            //                        Paragraphs=new System.Collections.Generic.List<Office.Model.AlizhouParagraph>(){
            //                            new Office.Model.AlizhouParagraph{
            //                                Run=new Office.Model.AlizhouRun{
            //                                    Text="周黎"
            //                                }
            //                            }
            //                        }

            //                    },
            //                    new Office.Model.AlizhouTableCell{
            //                        Width=100,
            //                        Paragraphs=new System.Collections.Generic.List<Office.Model.AlizhouParagraph>(){
            //                            new Office.Model.AlizhouParagraph{
            //                                Run=new Office.Model.AlizhouRun{
            //                                    Text="18"
            //                                }
            //                            }
            //                        }

            //                    }
            //                },
            //            },
            //               new Office.Model.AlizhouTableRow{
            //                Height=200,
            //                Cells=new System.Collections.Generic.List<Office.Model.AlizhouTableCell>(){
            //                    new Office.Model.AlizhouTableCell{
            //                        Width=100,
            //                        Paragraphs=new System.Collections.Generic.List<Office.Model.AlizhouParagraph>(){
            //                            new Office.Model.AlizhouParagraph{
            //                                Run=new Office.Model.AlizhouRun{
            //                                    Text="张三",
            //                                    IsBold=true,
            //                                    Pictures=new System.Collections.Generic.List<Office.Model.AlizhouPicture>()                     {
            //                                        new Office.Model.AlizhouPicture{ PictureUrl="D://191cb437-2bc8-4fe9-9c5c-b7536eae1883.jpg",Width=30,Height=30},
            //                                        new Office.Model.AlizhouPicture{ PictureUrl="D://renwu-mayun1.jpg",Width=30,Height=30}
            //                                    }
            //                                }
            //                            }
            //                        }

            //                    },
            //                    new Office.Model.AlizhouTableCell{
            //                        Width=100,
            //                        Paragraphs=new System.Collections.Generic.List<Office.Model.AlizhouParagraph>(){
            //                            new Office.Model.AlizhouParagraph{
            //                                Run=new Office.Model.AlizhouRun{
            //                                    Text="19",
            //                                    Color=Color.Red,
            //                                    FontFamily="微软雅黑",
            //                                    FontSize=12,
            //                                    IsBold=true,
            //                                }
            //                            }
            //                        }

            //                    }
            //                },
            //            }
            //        }
            //    }
            //};
            //var word = wordExportService.TemplateCreateWord(templateUrl, userTemplate);
            //File.WriteAllBytes($"{basePath}/{DateTime.Now.ToString("yyyyMMddHHmmss")}测试生成word.docx", word.WordBytes);



            //using (DocX doc = DocX.Create("D://测试测试.docx"))
            //{
            //    var table = doc.AddTable(4, 3);
            //    table.Alignment = Alignment.center;
            //    table.SetBorder(TableBorderType.InsideH, new Border { });
            //    table.SetBorder(TableBorderType.Top, new Border { });
            //    table.SetBorder(TableBorderType.Bottom, new Border { });
            //    table.SetBorder(TableBorderType.Left, new Border { });
            //    table.SetBorder(TableBorderType.Right, new Border { });
            //    table.SetBorder(TableBorderType.InsideV, new Border { });
            //    table.Rows[0].MergeCells(0, 3);
            //    table.Rows[0].Cells[0].Paragraphs[0].Append("质量风险评分汇总");
            //    table.Rows[1].Cells[0].Paragraphs[0].Append("分项工程");
            //    table.Rows[1].Cells[1].Paragraphs[0].Append("分项合格率");
            //    table.Rows[1].Cells[2].Paragraphs[0].Append("质量风险评估结果");
            //    table.Rows[2].Cells[0].Paragraphs[0].Append("渗漏");
            //    table.Rows[3].Cells[0].Paragraphs[0].Append("备注");
            //    table.Rows[3].Cells[1].Paragraphs[0].Append("质量风险评估结果 = 实得分 / 应得分 * 100 %。");

            //    doc.InsertTable(table);
            //    doc.Save();
            //}

            Console.WriteLine("\tMergeCells()");

            // Create a document.
            using (var document = DocX.Create("D://MergeCells.docx"))
            {
                // Add a title.
                document.InsertParagraph("Merge and delete cells").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

                // Add A table.
                var t = document.AddTable(3, 2);

                var t1 = document.InsertTable(t);

                // Add 4 columns in the table.
                t1.InsertColumn();
                t1.InsertColumn();
                t1.InsertColumn(t1.ColumnCount - 1, true);
                t1.InsertColumn(t1.ColumnCount - 1, true);

                // Merged Cells 1 to 4 in first row of the table.
                t1.Rows[0].MergeCells(1, 4);

                // Merged the last 2 Cells in the second row of the table.
                var columnCount = t1.Rows[1].ColumnCount;
                t1.Rows[1].MergeCells(columnCount - 2, columnCount - 1);

                // Add text in each cell of the table.
                foreach (var r in t1.Rows)
                {
                    for (int i = 0; i < r.Cells.Count; ++i)
                    {
                        var c = r.Cells[i];
                        c.Paragraphs[0].InsertText("Column " + i);
                        c.Paragraphs[0].Alignment = Alignment.center;
                    }
                }
                // Delete the second cell from the third row and shift the cells on its right by 1 to the left.
               // t1.DeleteAndShiftCellsLeft(2, 1);

                document.Save();
                Console.WriteLine("\tCreated: MergeCells.docx\n");
            }
            Console.WriteLine("yes");
            Console.ReadKey();
        }
    }
}
