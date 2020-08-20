using System;
using System.Collections.Generic;
using System.Text;

namespace Alizhou.Office.Attribute
{
    /// <summary>
    /// 普通占位符特性
    /// </summary>
    public class PlaceholderAttribute : System.Attribute
    {
        public PlaceholderAttribute(string placeHolder)
        {
            Placeholder = placeHolder;
        }
        public string Placeholder { get; set; }
    }
}
