using System;
using System.IO;
using Pchp.Core;
using Pchp.Library;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace nulastudio.Document.EPPlus4PHP
{
    public class ExcelConvert
    {
        public static int toIndex(string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+"))
            {
                throw new ArgumentException();
            }
            int index = 0;
            char[] chars = columnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            // ONE-BASE to ZERO-BASE
            // return index - 1;
            // ONE-BASE
            return index;
        }

        public static string toName(int index)
        {
            // ONE-BASE to ZERO-BASE
            index--;
            if (index < 0)
            {
                throw new ArgumentException();
            }
            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);
            return String.Join(string.Empty, chars.ToArray());
        }
    }
}
