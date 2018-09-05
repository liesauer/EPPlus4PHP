using System;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;

namespace nulastudio.Document.EPPlus4PHP
{
    public class WorkBook
    {
        private ExcelWorkbook _workBook;
        private WorkSheets _workSheets;
        private bool _is1Base;
        public WorkBook(ExcelWorkbook workBook, bool is1Base)
        {
            _workBook = workBook;
            _workSheets = new WorkSheets(_workBook.Worksheets, is1Base);
            _is1Base = is1Base;
        }
        public WorkSheets workSheets { get => _workSheets; }
        public bool is1Base { get => _is1Base; }
    }
}