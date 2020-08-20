using Alizhou.Office.Interfaces;
using Alizhou.Office.Model;
using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace Alizhou.Office.Helper
{
    public class DocXHelper
    {
        /// <summary>
        /// 获取模板文件
        /// </summary>
        /// <param name="fileUrl"></param>
        /// <returns></returns>
        public static DocX GetDocX
           (string fileUrl)
        {
            DocX word;

            if (!File.Exists(fileUrl))
            {
                throw new Exception("找不到模板文件");
            }

            try
            {
                using (FileStream fs = File.OpenRead(fileUrl))
                {
                    word = DocX.Load(fs);
                }
            }
            catch (Exception)
            {
                throw new Exception("打开模板文件失败");
            }

            return word;
        }
        /// <summary>
        /// 替换Word中的占位符
        /// </summary>
        /// <param name="word"></param>
        /// <param name="placeholderEntities">每个属性封装后的实体</param>
        public static void ReplacePlaceholdersInWord(DocX word, IEnumerable<PlaceholderEntity> placeholderEntities)
        {
            if (word == null)
                throw new ArgumentNullException("word");
            if (placeholderEntities == null)
                throw new ArgumentNullException("placeholderEntities");
            foreach (var placeholder in placeholderEntities)
            {
                switch (placeholder.PlaceholderType)
                {
                    case Enum.PlaceholderType.Table:
                        ReplacePlaceholdersInTable(word, placeholder.Placeholder, (AlizhouTable)placeholder.Data);
                        break;
                    case Enum.PlaceholderType.Text:
                        ReplacePlaceholdersInText(word, placeholder.Placeholder, (AlizhouText)placeholder.Data);
                        break;
                    case Enum.PlaceholderType.Picture:
                        ReplacePlaceholdersInImage(word, placeholder.Placeholder, placeholder.Pictures);
                        break;
                    default:
                        break;
                }

            }
        }
        private static void ReplacePlaceholdersInText(DocX word, string oldText, AlizhouText newText)
        {
            foreach (var paragraph in word.Paragraphs)
            {
                if (paragraph.Text.Contains(oldText))
                    paragraph.ReplaceText(oldText, newText.Data);
            }
        }
        private static void ReplacePlaceholdersInTable(DocX word, string oldText, AlizhouTable newTable)
        {
            foreach (var paragraph in word.Paragraphs)
            {
                if (paragraph.Text.Contains(oldText))
                {
                    var table = AlizhouTableToTable(word, newTable);
                    paragraph.InsertTableAfterSelf(table);
                    paragraph.Remove(false);
                }
            }
        }
        private static void ReplacePlaceholdersInImage(DocX word, string oldText, IEnumerable<AlizhouPicture> newPic)
        {
            if (newPic.Count() > 0)
            {
                foreach (var paragraph in word.Paragraphs)
                {
                    if (paragraph.Text.Contains(oldText))
                    {
                        var pics = newPic.ToList();
                        pics.ForEach(pic =>
                        {
                            Stream stream = pic.PictureData != null ? pic.PictureData : File.OpenRead(pic.PictureUrl);
                            paragraph.AppendPicture(word.AddImage(stream).CreatePicture(pic.Height, pic.Width));
                            paragraph.ReplaceText(oldText, "");
                        });

                    }
                }
            }
        }
        private static Table AlizhouTableToTable(DocX word, AlizhouTable alizhouTable)
        {
            var table = word.AddTable(alizhouTable.RowCount, alizhouTable.ColumnCount);
            table.Alignment = Alignment.center;
            table.SetBorder(TableBorderType.InsideH, new Border { });
            table.SetBorder(TableBorderType.Top, new Border { });
            table.SetBorder(TableBorderType.Bottom, new Border { });
            table.SetBorder(TableBorderType.Left, new Border { });
            table.SetBorder(TableBorderType.Right, new Border { });
            table.SetBorder(TableBorderType.InsideV, new Border { });
            for (int i = 0; i < alizhouTable.Rows.Count; i++)
            {
                table.Rows[i].Height = alizhouTable.Rows[i].Height;//设置行高
                //处理每行单元格
                for (int j = 0; j < alizhouTable.Rows[i].Cells.Count; j++)
                {
                    var alizhouTableCell = alizhouTable.Rows[i].Cells[j];
                    //设置单元格宽
                    table.Rows[i].Cells[j].Width = alizhouTableCell.Width;
                    if (alizhouTableCell.FillColor != Color.Empty)
                        table.Rows[i].Cells[j].FillColor = alizhouTableCell.FillColor;
                    foreach (var item in alizhouTableCell.Paragraphs)
                    {
                        if (item.Run.Pictures.Count > 0)
                        {
                            Paragraph paragraph = table.Rows[i].Cells[j].InsertParagraph();
                            paragraph.Alignment = item.Alignment;
                            item.Run.Pictures.ForEach(t =>
                            {
                                Stream stream = t.PictureData != null ? t.PictureData : File.OpenRead(t.PictureUrl);
                                paragraph.InsertPicture(word.AddImage(stream).CreatePicture(t.Width, t.Height));
                            });
                        }
                        else
                        {
                            //Formatting formatting = new Formatting
                            //{
                            //    Bold = item.Run.IsBold,
                            //    FontColor = item.Run.Color,
                            //    FontFamily = new Font(item.Run.FontFamily),
                            //    Size = item.Run.FontSize
                            //};

                            table.Rows[i].Cells[j].Paragraphs[0].Append(item.Run.Text);
                            if (item.Run.IsBold)
                                table.Rows[i].Cells[j].Paragraphs[0].Bold();
                            table.Rows[i].Cells[j].Paragraphs[0].FontSize(item.Run.FontSize);
                            table.Rows[i].Cells[j].Paragraphs[0].Font(item.Run.FontFamily);
                            table.Rows[i].Cells[j].Paragraphs[0].Color(item.Run.Color);
                            table.Rows[i].Cells[j].Paragraphs[0].Alignment = item.Alignment;
                        }
                    }
                }
            }
            return table;

        }
    }
}
