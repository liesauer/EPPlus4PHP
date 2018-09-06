using System;
using System.Collections.Generic;
using System.Linq;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;
using System.Text.RegularExpressions;

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

        #region Addressing
        // 单选 ["B1"]
        // 列选 ["A"] ["A:A"]
        // 行选 [1]["1"]["1:1"]
        // 窗选 ["A1:B2"]
        // 多选 ["A1:B2,A,8"]["A1:B2,A:A,1:1000"]
        public static string parseAddress(Context ctx, PhpValue address)
        {
            return parseAddress(address.ToString(ctx));
        }
        public static string parseAddress(Context ctx, PhpString address)
        {
            return parseAddress(address.ToString(ctx));
        }
        public static string parseAddress(string address)
        {
            #warning invalid address or out-of-bounds address cannot be detected till now.
            string normalizeAddress(string _address)
            {
                int row = 0;
                if (int.TryParse(_address, out row))
                {
                    // row
                    return $"{row}:{row}";
                }
                else
                {
                    // pure column
                    if (Regex.IsMatch(_address, "^[A-Z]+$"))
                    {
                        return $"{_address}:{_address}";
                    }
                    else
                    {
                        return _address;
                    }
                }
            }
            address = Regex.Replace(address, "[^0-9a-zA-Z\\:\\,]", "").ToUpper();
            string[] addresses = address.Split(',');
            List<string> res = new List<string>();
            foreach (string addr in addresses)
            {
                res.Add(normalizeAddress(addr));
            }
            return string.Join(",",res);
        }

        #endregion



    }
}