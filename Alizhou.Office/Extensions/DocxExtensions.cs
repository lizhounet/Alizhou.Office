using Novacode;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Alizhou.Office.Extensions
{
    public static class DocxExtensions
    {
        public static byte[] ToBytes(this DocX doc)
        {
            byte[] result;
            using (MemoryStream ms = new MemoryStream())
            {
                doc.SaveAs(ms);
                result = ms.ToArray();
            }
            return result;
        }
    }
}
