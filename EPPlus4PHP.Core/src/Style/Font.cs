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
        public Color color
        {
            get
            {
                string hexARGB = _font.Color.Rgb;
                string hexA = hexARGB.Substring(0, 2);
                string hexR = hexARGB.Substring(2, 2);
                string hexG = hexARGB.Substring(4, 2);
                string hexB = hexARGB.Substring(6, 2);
                int a = 0, r = 0, g = 0, b = 0;
                a = int.Parse(hexA, System.Globalization.NumberStyles.HexNumber);
                r = int.Parse(hexR, System.Globalization.NumberStyles.HexNumber);
                g = int.Parse(hexG, System.Globalization.NumberStyles.HexNumber);
                b = int.Parse(hexB, System.Globalization.NumberStyles.HexNumber);
                Color color = new Color(a, r, g, b);
                // will cause var_dump throwing StackOverflowException
                // color.ValueChanged += ColorChanged;
                return color;
            }
            set
            {
                Color color = value as Color;
                _font.Color.SetColor(color.alpha, color.red, color.green, color.blue);
            }
        }
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

        internal void ColorChanged(object sender, EventArgs e)
        {
            Color color = sender as Color;
            _font.Color.SetColor(color.alpha, color.red, color.green, color.blue);
        }
    }
}