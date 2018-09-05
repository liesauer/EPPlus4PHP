using System;
using System.Collections.Generic;
using Pchp.Core;
using Pchp.Library;
using Pchp.Library.Spl;
using OfficeOpenXml;
using System.Collections;

namespace nulastudio.Document.EPPlus4PHP
{
    public class WorkSheets : ArrayAccess, Countable, IEnumerable<KeyValuePair<string, WorkSheet>>, IEnumerator<KeyValuePair<string, WorkSheet>>, IDisposable
    {
        private ExcelWorksheets _workSheets;
        private List<KeyValuePair<string, WorkSheet>> _workSheetsList;
        private int _count = 0;
        private int _pointer = 0;
        private bool _is1Base;
        public WorkSheets(ExcelWorksheets workSheets, bool is1Base)
        {
            _workSheets = workSheets;
            _workSheetsList = new List<KeyValuePair<string, WorkSheet>>();
            _is1Base = is1Base;
        }

        public bool is1Base { get => _is1Base; }

        KeyValuePair<string, WorkSheet> IEnumerator<KeyValuePair<string, WorkSheet>>.Current => _workSheetsList[_pointer - 1];

        object IEnumerator.Current => _workSheetsList[_pointer - 1].Value;

        #region ArrayAccess
        public PhpValue offsetGet(PhpValue offset)
        {
            if (offsetExists(offset))
            {
                IntStringKey key = default(IntStringKey);
                if (offset.TryToIntStringKey(out key))
                {
                    ExcelWorksheet excelWorksheet;
                    if (key.IsInteger)
                    {
                        int index = key.Integer;
                        // 需要兼容ZERO-BASE以及ONE-BASE
                        if (!is1Base)
                        {
                            index--;
                        }
                        excelWorksheet = _workSheets[index];
                    }
                    else
                    {
                        excelWorksheet = _workSheets[key.String];
                    }
                    return PhpValue.FromClr(new WorkSheet(excelWorksheet, is1Base));
                }
            } else {
                // 不存在时则创建（使用表名时，使用数字索引不会创建）
                IntStringKey key = default(IntStringKey);
                if (offset.TryToIntStringKey(out key) && !string.IsNullOrEmpty(key.String))
                {
                    add(key.String);
                    return offsetGet(offset);
                }
            }
            return PhpValue.Null;
        }
        public void offsetSet(PhpValue offset, PhpValue value)
        {
            throw new NotImplementedException();
        }
        public void offsetUnset(PhpValue offset)
        {
            IntStringKey key = default(IntStringKey);
            if (offset.TryToIntStringKey(out key))
            {
                if (key.IsInteger)
                {
                    int index = key.Integer;
                    // 需要兼容ZERO-BASE以及ONE-BASE
                    if (!is1Base)
                    {
                        index--;
                    }
                    _workSheets.Delete(index);
                } else {
                    _workSheets.Delete(key.String);
                }
            }
            else
            {
                throw new NotSupportedException();
            }
        }
        public bool offsetExists(PhpValue offset)
        {
            try
            {
                IntStringKey key = default(IntStringKey);
                if (offset.TryToIntStringKey(out key))
                {
                    if (key.IsInteger)
                    {
                        // 需要兼容ZERO-BASE以及ONE-BASE
                        #warning 不检查索引会导致EPPlus抛出System.IndexOutOfRangeException异常
                        return (is1Base ? _workSheets[key.Integer] : _workSheets[key.Integer - 1]) != null;
                    }
                    else
                    {
                        return _workSheets[key.String] != null;
                    }
                }
                return false;
            }
            catch (System.Exception)
            {
                return false;
            }
        }
        #endregion

        #region Countable
        public long count()
        {
            return _workSheets.Count;
        }
        #endregion

        #region IEnumerable
        IEnumerator<KeyValuePair<string, WorkSheet>> IEnumerable<KeyValuePair<string, WorkSheet>>.GetEnumerator()
        {
            updateDic();
            return this as IEnumerator<KeyValuePair<string, WorkSheet>>;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            updateDic();
            return this as IEnumerator;
        }
        private void updateDic()
        {
            _workSheetsList.Clear();
            _count = 0;
            _pointer = 0;
            foreach (ExcelWorksheet worksheet in _workSheets)
            {
                _count++;
                _workSheetsList.Add(new KeyValuePair<string, WorkSheet>(worksheet.Name, new WorkSheet(worksheet, is1Base)));
            }
        }
        #endregion

        #region IEnumerator
        bool IEnumerator.MoveNext()
        {
            return _count != 0 && ++_pointer <= _count;
        }

        void IEnumerator.Reset()
        {
            _pointer = 0;
        }

        void IDisposable.Dispose()
        {
            _workSheetsList.Clear();
        }
        #endregion
    
        #region Add WorkSheet
        public void add(Context ctx, PhpString sheetName)
        {
            add(ctx, sheetName.ToString(ctx));
        }
        public void add(Context ctx, string sheetName)
        {
            add(sheetName);
        }
        public void add(string sheetName)
        {
            _workSheets.Add(sheetName);
        }
        #endregion

        #region Delete WorkSheet
        public void delete(Context ctx, PhpString sheetName)
        {
            delete(ctx, sheetName.ToString(ctx));
        }
        public void delete(Context ctx, string sheetName)
        {
            delete(sheetName);
        }
        public void delete(string sheetName)
        {
            _workSheets.Delete(sheetName);
        }
        #endregion

    }
}