using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml.Style;
using EPBorder = OfficeOpenXml.Style.Border;

namespace nulastudio.Document.EPPlus4PHP.Style
{
    public class Border
    {
        private EPBorder _border;
        private BorderItem _top;
        private BorderItem _bottom;
        private BorderItem _left;
        private BorderItem _right;
        private BorderItem _diagonal;
        // private BorderItem _horizontalMiddle;
        // private BorderItem _verticalMiddle;

        public Border(EPBorder border)
        {
            _border = border;
            _top = new BorderItem(_border.Top);
            _bottom = new BorderItem(_border.Bottom);
            _left = new BorderItem(_border.Left);
            _right = new BorderItem(_border.Right);
            _diagonal = new BorderItem(_border.Diagonal);
        }

        public BorderItem top { get => _top; }
        public BorderItem bottom { get => _bottom; }
        public BorderItem left { get => _left; }
        public BorderItem right { get => _right; }
        public BorderItem diagonal { get => _diagonal; }

        // left to right
        public bool diagonalUp { get => _border.DiagonalUp; set => _border.DiagonalUp = value; }
        public bool diagonalDown { get => _border.DiagonalDown; set => _border.DiagonalDown = value; }
    }
}