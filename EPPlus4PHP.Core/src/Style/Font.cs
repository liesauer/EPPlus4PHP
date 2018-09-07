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
        private Color _fontColor;
        
        public Font(ExcelFont font)
        {
            _font = font;
            string hexARGB = _font.Color.Rgb;
            if (string.IsNullOrEmpty(hexARGB))
            {
                _fontColor = (Color)Color.BLACK_COLOR;
            } else {
                string hexA = hexARGB.Substring(0, 2);
                string hexR = hexARGB.Substring(2, 2);
                string hexG = hexARGB.Substring(4, 2);
                string hexB = hexARGB.Substring(6, 2);
                int a = 0, r = 0, g = 0, b = 0;
                a = int.Parse(hexA, System.Globalization.NumberStyles.HexNumber);
                r = int.Parse(hexR, System.Globalization.NumberStyles.HexNumber);
                g = int.Parse(hexG, System.Globalization.NumberStyles.HexNumber);
                b = int.Parse(hexB, System.Globalization.NumberStyles.HexNumber);
                _fontColor = new Color(a, r, g, b);
            }
            // will cause var_dump throwing StackOverflowException
            _fontColor.ValueChanged += ColorChanged;
        }

        public string name { get => _font.Name; set => _font.Name = value; }
        public float size { get => _font.Size; set => _font.Size = value; }
        public int family { get => _font.Family; set => _font.Family = value; }
        public Color color
        {
            get => _fontColor;
            set
            {
                Color color = value as Color;
                _fontColor.setColor(color.alpha, color.red, color.green, color.blue);
                _font.Color.SetColor(color.alpha, color.red, color.green, color.blue);
            }
        }
        public string scheme { get => _font.Scheme; set => _font.Scheme = value; }
        public bool bold { get => _font.Bold; set => _font.Bold = value; }
        public bool italic { get => _font.Italic; set => _font.Italic = value; }
        public bool strike { get => _font.Strike; set => _font.Strike = value; }
        // if set to false, underLineType will be set to None
        // if set to true, underLineType will be set to Single
        public bool underLine { get => _font.UnderLine; set => _font.UnderLine = value; }
        public int underLineType { get => (int)_font.UnderLineType; set => _font.UnderLineType = (ExcelUnderLineType)value; }
        # warning BUG: effective only for the first letter
        public int verticalAlign { get => (int)_font.VerticalAlign; set => _font.VerticalAlign = (ExcelVerticalAlignmentFont)value; }

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

        internal void ColorChanged(object sender, EventArgs e)
        {
            this.color = sender as Color;
        }
    }
}