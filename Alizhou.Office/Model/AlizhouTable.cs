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
            RowCount = rowCount;
            ColumnCount = columnCount;
        }
        public List<AlizhouTableRow> Rows { get; set; }
        /// <summary>
        /// 总行数
        /// </summary>
        public int RowCount { set; get; }
        /// <summary>
        /// 总列数
        /// </summary>
        public int ColumnCount { set; get; }

    }

}
