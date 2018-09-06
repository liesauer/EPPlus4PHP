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
        
        public Style(ExcelStyle style)
        {
            _style = style;
            _font = new Font(style.Font);
        }

        public Font font { get => _font; }
    }
}