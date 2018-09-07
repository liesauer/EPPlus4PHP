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
        public const long BLACK_COLOR = 0xFF000000;      // 0.0 white
        public const long DARKGRAY_COLOR = 0xFF555555;   // 0.333 white
        public const long LIGHTGRAY_COLOR = 0xFFAAAAAA;  // 0.667 white
        public const long WHITE_COLOR = 0xFFFFFFFF;      // 1.0 white
        public const long GRAY_COLOR = 0xFF7F7F7F;       // 0.5 white
        public const long RED_COLOR = 0xFFFF0000;        // 1.0, 0.0, 0.0 RGB
        public const long GREEN_COLOR = 0xFF00FF00;      // 0.0, 1.0, 0.0 RGB
        public const long BLUE_COLOR = 0xFF0000FF;       // 0.0, 0.0, 1.0 RGB
        public const long CYAN_COLOR = 0xFF00FFFF;       // 0.0, 1.0, 1.0 RGB
        public const long YELLOW_COLOR = 0xFFFFFF00;     // 1.0, 1.0, 0.0 RGB
        public const long MAGENTA_COLOR = 0xFFFF00FF;    // 1.0, 0.0, 1.0 RGB
        public const long ORANGE_COLOR = 0xFFFF7F00;     // 1.0, 0.5, 0.0 RGB
        public const long PURPLE_COLOR = 0xFF7F007F;     // 0.5, 0.0, 0.5 RGB
        public const long BROWN_COLOR = 0xFF996633;      // 0.6, 0.4, 0.2 RGB
        public const long CLEAR_COLOR = 0x00000000;      // 0.0 white, 0.0 alpha
        private int _alpha;
        private int _red;
        private int _green;
        private int _blue;

        public event EventHandler<EventArgs> ValueChanged;

        public Color() : this(0, 255, 255, 255)
        {
        }
        public Color(int alpha, int red, int green, int blue)
        {
            _alpha = alpha;
            _red = red;
            _green = green;
            _blue = blue;
        }

        public int alpha
        {
            get => _alpha;
            set
            {
                _alpha = value & 0xFF;
                OnValueChanged(new EventArgs());
            }
        }
        public int red
        {
            get => _red;
            set
            {
                _red = value & 0xFF;
                OnValueChanged(new EventArgs());
            }
        }
        public int green
        {
            get => _green;
            set
            {
                _green = value & 0xFF;
                OnValueChanged(new EventArgs());
            }
        }
        public int blue
        {
            get => _blue;
            set
            {
                _blue = value & 0xFF;
                OnValueChanged(new EventArgs());
            }
        }

        public static implicit operator Color(long aRBG)
        {
            int alpha = (int)((aRBG & 0xFF000000) >> 24);
            int red = (int)((aRBG & 0xFF0000) >> 16);
            int green = (int)((aRBG & 0xFF00) >> 8);
            int blue = (int)(aRBG & 0xFF);
            return new Color(alpha, red, green, blue);
        }
        public static implicit operator long(Color color)
        {
            return color.alpha << 24 & color.red << 16 & color.green << 8 & color.blue;
        }

        protected virtual void OnValueChanged(EventArgs e)
        {
            if (ValueChanged != null)
            {
                ValueChanged(this, e);
            }
        }
    }
}