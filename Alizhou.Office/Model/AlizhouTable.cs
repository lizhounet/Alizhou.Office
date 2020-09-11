using Alizhou.Office.Interfaces;
using Novacode;
using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Model
{
    public class AlizhouTable : IWordElement
    {
        public AlizhouTable() { }
        public AlizhouTable(int rowCount, int columnCount)
        {
            ColumnCount = columnCount;
            for (int i = 0; i < rowCount; i++)
            {
                var row = new AlizhouTableRow();
                for (int j = 0; j < columnCount; j++)
                {
                    var cell = new AlizhouTableCell
                    {
                        Paragraphs = new List<AlizhouParagraph>() {
                        new AlizhouParagraph
                        {
                            Run=new AlizhouRun()
                        }
                        }
                    };
                    row.Cells.Add(cell);
                }
                this.Rows.Add(row);
            }
        }
        public List<AlizhouTableRow> Rows { get; set; } = new List<AlizhouTableRow>();
        /// <summary>
        /// 总行数
        /// </summary>
        public int RowCount { get { return this.Rows.Count; } }
        /// <summary>
        /// 总列数
        /// </summary>
        public int ColumnCount { get; }
        /// <summary>
        /// 需要合并的列单元格
        /// </summary>
        internal List<(int, int, int)> MergeCellsInColumns { get; } = new List<(int, int, int)>();
        /// <summary>
        /// 需要合并的行单元格
        /// </summary>
        internal List<(int, int, int)> MergeCellsInRows { get; } = new List<(int, int, int)>();
        /// <summary>
        /// 合并列单元格
        /// </summary>
        /// <param name="columnIndex">合并第几列</param>
        /// <param name="startRow">合并开始行</param>
        /// <param name="endRow">合并结束第几行</param>
        public void MergeCellsInColumn(int columnIndex, int startRow, int endRow)
        {
            MergeCellsInColumns.Add((columnIndex, startRow, endRow));
        }
        /// <summary>
        /// 合并行单元格
        /// </summary>
        /// <param name="rowIndex">合并第几行</param>
        /// <param name="startColumn">合并开始列</param>
        /// <param name="endColumn">合并结束第几列</param>
        public void MergeCellsInRow(int rowIndex, int startColumn, int endColumn)
        {
            MergeCellsInRows.Add((rowIndex, startColumn, endColumn));
        }
    }

}
