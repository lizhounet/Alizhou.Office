using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Model
{
    public class AlizhouTableRow
    {
        public List<AlizhouTableCell> Cells { get; set; } = new List<AlizhouTableCell>();
        /// <summary>
        /// 行高
        /// </summary>
        public double Height { get; set; }
    }
}
