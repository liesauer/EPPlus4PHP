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
        // [A-Z]+[1-9][0-9]*:[A-Z]+[1-9][0-9]*
        public const string REGEX_PRUE_ROW = "^[1-9][0-9]*$";
        public const string REGEX_PRUE_COLUMN = "^[A-Z]+$";
        public const string REGEX_SINGLE_CELL = "^[A-Z]+[1-9][0-9]*$";
        public const string REGEX_SINGLE_ROW = "^[1-9][0-9]*:[1-9][0-9]*$";
        public const string REGEX_SINGLE_COLUMN = "^[A-Z]+:[A-Z]+$";
        public const string REGEX_MULTI_CELLS = "^^[A-Z]+[1-9][0-9]*:[A-Z]+[1-9][0-9]*$";

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
        public string address { get => _range.Address; }
        public string fullAddress { get => _range.FullAddress; }
        public string fullAddressAbsolute { get => _range.FullAddressAbsolute; }
        public string from { get => _range.Start.Address; }
        public int fromRow { get => _range.Start.Row; }
        public int fromColumn { get => _range.Start.Column; }
        public string to { get => _range.End.Address; }
        public int toRow { get => _range.End.Row; }
        public int toColumn { get => _range.End.Column; }
        public int rows { get => _range.Rows; }
        public int columns { get => _range.Columns; }
        public PhpValue value
        {
            get
            {
                if (_range.Value is object[,])
                {
                    object[,] data = _range.Value as object[,];
                    int rows = data.GetLength(0);
                    PhpArray arr = new PhpArray();
                    for (int i = 0; i < rows; i++)
                    {
                        PhpArray rowData = new PhpArray();
                        int columns = data.GetLength(1);
                        for (int j = 0; j < columns; j++)
                        {
                            rowData.AddValue(PhpValue.FromClr(data.GetValue(i,j)));
                        }
                        arr.AddValue(PhpValue.Create(rowData));
                    }
                    return arr;
                }
                else
                {
                    return PhpValue.FromClr(_range.Value);
                }
            }
            set
            {
                if (value.IsArray)
                {
                    List<List<object>> data = new List<List<object>>();
                    foreach (KeyValuePair<IntStringKey, PhpValue> item in value.ToArray())
                    {
                        if (item.Value.IsArray)
                        {
                            List<object> rowData = new List<object>();
                            foreach (KeyValuePair<IntStringKey, PhpValue> cell in item.Value.ToArray())
                            {
                                rowData.Add(cell.Value.ToClr());
                            }
                            data.Add(rowData);
                        } else {
                            data.Add(new List<object>(){item.Value.ToClr()});
                        }
                    }

                    int fromRow = this.fromRow;
                    int fromColumn = this.fromColumn;
                    int row = Math.Min(rows, data.Count);
                    int column = Math.Min(columns, data.Count == 0 ? 0 : data[0].Count);
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < column; j++)
                        {
                            _range[fromRow + i, fromColumn + j].Value = data[i][j];
                        }
                    }

                    ;
                }
                else
                {
                    _range.Value = value.ToClr();
                }
            }
        }
        public bool merge
        {
            get => _range.Merge;
            set
            {
                try
                {
                    _range.Merge = value;
                }
                catch {}
            }
        }

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
                string address = key.Object.ToString();
                if (tryParseAddress(address, out address))
                {
                    return PhpValue.FromClr(this[address]);
                }
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
        // public static bool tryParseAddress(Context ctx, PhpValue address, PhpAlias addr_out)
        // {
        //     bool res = tryParseAddress(address.ToString(ctx), out var tmp_addr);
        //     if (res)
        //     {
        //         addr_out.Value = PhpValue.Create(tmp_addr);
        //     }
        //     return res;
        // }
        // public static bool tryParseAddress(Context ctx, PhpString address, PhpAlias addr_out)
        // {
        //     bool res = tryParseAddress(address.ToString(ctx), out var tmp_addr);
        //     if (res)
        //     {
        //         addr_out.Value = PhpValue.Create(tmp_addr);
        //     }
        //     return res;
        // }
        public static bool tryParseAddress(string address, out string address_out)
        {
            bool tryNormalizeAddress(string addr_in, out string addr_out)
            {
                if (int.TryParse(addr_in, out var row) &&
                    row >= 1 &&
                    row <= ExcelPackage.MAX_ROWS)
                {
                    // pure row
                    addr_out = $"{row}:{row}";
                    return true;
                }
                else if (Regex.IsMatch(addr_in, REGEX_PRUE_COLUMN) &&
                         ExcelConvert.toIndex(addr_in) <= ExcelPackage.MAX_COLUMNS)
                {
                    // pure column
                    addr_out = $"{addr_in}:{addr_in}";
                    return true;
                }
                else if (Regex.IsMatch(addr_in, REGEX_SINGLE_CELL))
                {
                    // single cell
                    string addr1 = Regex.Match(addr_in, @"^[A-Z]+").Value;
                    string addr2 = Regex.Match(addr_in, @"[1-9][0-9]*$").Value;
                    if (tryNormalizeAddress(addr1, out var tmp_addr1) &&
                        tryNormalizeAddress(addr2, out var tmp_addr2))
                    {
                        addr_out = addr_in;
                        return true;
                    }
                }
                else if (Regex.IsMatch(addr_in, REGEX_SINGLE_ROW)    ||
                         Regex.IsMatch(addr_in, REGEX_SINGLE_COLUMN) ||
                         Regex.IsMatch(addr_in, REGEX_MULTI_CELLS))
                {
                    // row:row
                    // column:column
                    // multi cells
                    string[] addrs = addr_in.Split(':');
                    if (tryNormalizeAddress(addrs[0], out var tmp_addr1) &&
                        tryNormalizeAddress(addrs[1], out var tmp_addr2))
                    {
                        addr_out = addr_in;
                        return true;
                    }
                }
                addr_out = "";
                return false;
            }
            address = Regex.Replace(address, @"\s", "").ToUpper();
            List<string> res = new List<string>();
            foreach (string addr in address.Split(','))
            {
                if (tryNormalizeAddress(addr, out var addr_out))
                {
                    res.Add(addr_out);
                }
            }
            address_out = string.Join(",",res);
            return res.Count != 0;
        }
        #endregion



    }
}