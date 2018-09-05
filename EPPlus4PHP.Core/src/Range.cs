using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;

namespace nulastudio.Document.EPPlus4PHP
{
    public class Range : ArrayAccess
    {
        private ExcelRange _range;
        private Style.Style _style;
        private bool _is1Base;
        public Range(ExcelRange range, bool is1Base)
        {
            _range = range;
            _is1Base = is1Base;
            _style = new Style.Style(range.Style);
        }

        public bool is1Base { get => _is1Base; }
        public Style.Style style { get => _style; }

        #region Indexer
        public Range this[string address]
        {
            get => new Range(_range[address], is1Base);
        }
        #endregion


        #region ArrayAccess
        public PhpValue offsetGet(PhpValue offset)
        {
            IntStringKey key = default(IntStringKey);
            if (offset.TryToIntStringKey(out key))
            {
                return PhpValue.FromClr(this[key.String]);
            }
            return PhpValue.Null;
        }
        public void offsetSet(PhpValue offset, PhpValue value)
        {
            throw new NotSupportedException();
        }
        public void offsetUnset(PhpValue offset)
        {
            throw new NotSupportedException();
        }
        public bool offsetExists(PhpValue offset)
        {
            return false;
        }
        #endregion



    }
}