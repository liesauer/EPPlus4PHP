using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml.Style;

namespace nulastudio.Document.EPPlus4PHP.Style
{
    public class Font
    {
        private ExcelFont _font;
        
        public Font(ExcelFont font)
        {
            _font = font;
        }

        public string name { get => _font.Name; set => _font.Name = value; }
        public float size { get => _font.Size; set => _font.Size = value; }
        public int family { get => _font.Family; set => _font.Family = value; }


    }
}