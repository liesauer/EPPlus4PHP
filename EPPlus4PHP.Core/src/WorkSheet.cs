using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;

namespace nulastudio.Document.EPPlus4PHP
{
    public class WorkSheet
    {
        private ExcelWorksheet _workSheet;
        private bool _is1Base;
        public WorkSheet(ExcelWorksheet workSheet, bool is1Base)
        {
            _workSheet = workSheet;
            _is1Base = is1Base;
        }

        public bool is1Base { get => _is1Base; }
        public string name { get => _workSheet.Name; }
        public Range cells { get =>new Range(_workSheet.Cells, is1Base); }
        public bool hasData { get => _workSheet.Dimension != null; }
        public Range datas
        {
            get
            {
                if (hasData)
                {
                    return cells[_workSheet.Dimension.Address];
                }
                return null;
            }
        }

        #region Movement
        public void moveBefore(string targetName)
        {
            try
            {
                _workSheet.Workbook.Worksheets.MoveBefore(name, targetName);
            }
            catch {}
        }
        public void moveAfter(string targetName)
        {
            try
            {
                _workSheet.Workbook.Worksheets.MoveAfter(name, targetName);
            }
            catch {}
        }
        public void moveToStart()
        {
            try
            {
                _workSheet.Workbook.Worksheets.MoveToStart(name);
            }
            catch {}
        }
        public void moveToEnd(string sourceName)
        {
            try
            {
                _workSheet.Workbook.Worksheets.MoveToEnd(name);
            }
            catch {}
        }
        #endregion

        #region Cell RW
        public void addRow(Context ctx, PhpArray row)
        {
            int rowIndex = (hasData ? datas.toRow : 0) + 1;
            int startColumn = 1;
            foreach (PhpValue item in row.Values)
            {
                string columnName = ExcelConvert.toName(startColumn++);
                cells[string.Format("{0}{1}",columnName,rowIndex)].value = item;
            }
        }
        public void addColumn(Context ctx, PhpArray column)
        {
            int columnIndex = (hasData ? datas.toColumn : 0) + 1;
            string columnName = ExcelConvert.toName(columnIndex);
            int startRow = 1;
            foreach (PhpValue item in column.Values)
            {
                cells[string.Format("{0}{1}", columnName, startRow++)].value = item;
            }
        }
        #endregion
    }
}