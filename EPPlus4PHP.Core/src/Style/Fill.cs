using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml.Style;
using nulastudio.KVO;

namespace nulastudio.Document.EPPlus4PHP.Style
{
    public class Fill
    {
        private ExcelFill _fill;
        private Color _backgroundColor;

        public Fill(ExcelFill fill)
        {
            _fill = fill;
            string hexARGB = _fill.BackgroundColor.Rgb;
            if (string.IsNullOrEmpty(hexARGB))
            {
                _backgroundColor = (Color)Color.BLACK_COLOR;
            }
            else
            {
                string hexA = hexARGB.Substring(0, 2);
                string hexR = hexARGB.Substring(2, 2);
                string hexG = hexARGB.Substring(4, 2);
                string hexB = hexARGB.Substring(6, 2);
                int a = 0, r = 0, g = 0, b = 0;
                a = int.Parse(hexA, System.Globalization.NumberStyles.HexNumber);
                r = int.Parse(hexR, System.Globalization.NumberStyles.HexNumber);
                g = int.Parse(hexG, System.Globalization.NumberStyles.HexNumber);
                b = int.Parse(hexB, System.Globalization.NumberStyles.HexNumber);
                _backgroundColor = new Color(a, r, g, b);
            }
            #warning will cause var_dump throws StackOverflowException
            _backgroundColor.OnValueChanged += BackgroundColorChanged;
        }

        public Color backgroundColor
        {
            get => _backgroundColor;
            set
            {
                Color color = value as Color;
                _backgroundColor.setColor(color.alpha, color.red, color.green, color.blue);
                _fill.PatternType = ExcelFillStyle.Solid;
                _fill.BackgroundColor.SetColor(color.alpha, color.red, color.green, color.blue);
            }
        }
        internal void BackgroundColorChanged(object sender, ValueChangedEventArgs e)
        {
            this.backgroundColor = sender as Color;
        }
    }
}