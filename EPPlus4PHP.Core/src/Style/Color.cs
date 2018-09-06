using System;
using System.Drawing;
using System.Drawing.Text;
using SysFont = System.Drawing.Font;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml.Style;

namespace nulastudio.Document.EPPlus4PHP.Style
{
    public class Color
    {
        public const long RGB = 0xFF0000;
        public const long GREEN = 0x00FF00;
        public const long BLUE = 0x0000FF;
        private int _alpha;
        private int _red;
        private int _green;
        private int _blue;
        
        public Color() : this(0, 255, 255, 255)
        {
        }
        public Color(int alpha, int red, int green, int blue)
        {
            _alpha = alpha;
            _red = red;
            _green = green;
            _blue = green;
        }

        public int alpha { get => _alpha; set => _alpha = value & 0xFF; }
        public int red { get => _red; set => _red = value & 0xFF; }
        public int green { get => _green; set => _green = value & 0xFF; }
        public int blue { get => _blue; set => _blue = value & 0xFF; }

        public static implicit operator Color(long aRBG)
        {
            int alpha = (int)(aRBG & 0xFF000000) >> 24;
            int red = (int)(aRBG & 0xFF0000) >> 16;
            int green = (int)(aRBG & 0xFF00) >> 8;
            int blue = (int)(aRBG & 0xFF);
            return new Color(alpha, red, green, blue);
        }
        public static implicit operator long(Color color)
        {
            return color.alpha << 24 & color.red << 16 & color.green << 8 & color.blue;
        }
    }
}