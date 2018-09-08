using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml.Style;

namespace nulastudio.Document.EPPlus4PHP.Style
{
    public class Style
    {
        private ExcelStyle _style;
        private Font _font;
        private Fill _fill;
        
        public Style(ExcelStyle style)
        {
            _style = style;
            _font = new Font(style.Font);
            _fill = new Fill(style.Fill);
        }

        public Font font { get => _font; }
        public Fill fill { get => _fill; }
    }
}