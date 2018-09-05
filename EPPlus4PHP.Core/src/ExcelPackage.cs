using System;
using System.IO;
using Pchp.Core;
using Pchp.Library;
using EPExcelPackage = OfficeOpenXml.ExcelPackage;

namespace nulastudio.Document.EPPlus4PHP
{
    public class ExcelPackage
    {
        private EPExcelPackage _excelPackage;
        private WorkBook _workBook;
        public static readonly int MAX_ROWS = EPExcelPackage.MaxRows;
        public static readonly int MAX_COLUMNS = EPExcelPackage.MaxColumns;

        public ExcelPackage(Context ctx, PhpString fileName) : this(ctx, fileName.ToString(ctx))
        {
        }
        public ExcelPackage(Context ctx, string fileName)
        {
            _excelPackage = new EPExcelPackage(new FileInfo(fileName));
            _workBook = new WorkBook(_excelPackage.Workbook, _excelPackage.Compatibility.IsWorksheets1Based);
        }
        public static ExcelPackage open(Context ctx, PhpString fileName)
        {
            return open(ctx, fileName.ToString(ctx));
        }
        public static ExcelPackage open(Context ctx, string fileName)
        {
            return new ExcelPackage(ctx, fileName);
        }
        public void save(Context ctx)
        {
            // 空白工作簿将出错
            if (_excelPackage.Workbook.Worksheets.Count == 0)
            {
                _excelPackage.Workbook.Worksheets.Add("sheet1");
            }
            _excelPackage.Save();
        }

        public WorkBook workBook { get => _workBook; }
    }
}
