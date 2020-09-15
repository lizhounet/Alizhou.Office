using Alizhou.Office.Helper;
using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using Alizhou.Office.Provider;
using Alizhou.Office.Services;
using Alizhou.Test.Execl;
using Alizhou.Test.TestModel;
using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

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



            // EvaluationReportUniversal.EvaluationReportTemplateWrod();



            //EXECL
            ExeclImportExportTest.Test();

            Console.WriteLine("yes");
            Console.ReadKey();
        }
    }
}
