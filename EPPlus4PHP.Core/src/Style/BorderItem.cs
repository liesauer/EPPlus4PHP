using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml.Style;
using nulastudio.KVO;

namespace nulastudio.Document.EPPlus4PHP.Style
{
    public class BorderItem
    {
        private ExcelBorderItem _borderItem;
        private Color _color;

        public BorderItem(ExcelBorderItem borderItem)
        {
            _borderItem = borderItem;
            string hexARGB = _borderItem.Color.Rgb;
            if (string.IsNullOrEmpty(hexARGB))
            {
                _color = (Color)Color.BLACK_COLOR;
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
                _color = new Color(a, r, g, b);
            }
            #warning will cause var_dump throws StackOverflowException
            _color.OnValueChanged += ColorChanged;
        }

        public int style { get => (int)_borderItem.Style; set => _borderItem.Style = (ExcelBorderStyle)value; }
        public Color color
        {
            get => _color;
            set
            {
                Color color = value as Color;
                _color.setColor(color.alpha, color.red, color.green, color.blue);
                _borderItem.Color.SetColor(color.alpha, color.red, color.green, color.blue);
            }
        }

        internal void ColorChanged(object sender, ValueChangedEventArgs e)
        {
            this.color = sender as Color;
        }
    }
}