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
    }
}