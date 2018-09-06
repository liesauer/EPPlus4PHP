using System;
using System.Drawing;
using System.Drawing.Text;
using SysFont = System.Drawing.Font;
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
        public Color color { get; set; }
        public bool bold { get => _font.Bold; set => _font.Bold = value; }

        public bool tryLoadFont(string fontName, int fontSize = 12)
        {
            try
            {
                PrivateFontCollection prc = new PrivateFontCollection();
                prc.AddFontFile(fontName);
                SysFont font = new SysFont(prc.Families[0], fontSize);
                _font.SetFromFont(font);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}