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
        public ExcelPackage(Context ctx, PhpString fileName, PhpString password) : this(ctx, fileName.ToString(ctx), password.ToString(ctx))
        {
        }
        public ExcelPackage(Context ctx, string fileName)
        {
            _excelPackage = new EPExcelPackage(new FileInfo(fileName));
            _workBook = new WorkBook(_excelPackage.Workbook, _excelPackage.Compatibility.IsWorksheets1Based);
        }
        public ExcelPackage(Context ctx, string fileName, string password)
        {
            _excelPackage = new EPExcelPackage(new FileInfo(fileName), password);
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
        public static ExcelPackage open(Context ctx, PhpString fileName, PhpString password)
        {
            return open(ctx, fileName.ToString(ctx), password.ToString(ctx));
        }
        public static ExcelPackage open(Context ctx, string fileName, string password)
        {
            return new ExcelPackage(ctx, fileName, password);
        }
        public void save()
        {
            save(_excelPackage.Encryption.Password);
        }
        public void save(Context ctx, PhpString password)
        {
            save(password.ToString(ctx));
        }
        public void save(string password = null)
        {
            if (string.IsNullOrEmpty(password))
            {
                password = null;
            }
            checkWorkbookIsEmpty();
            _excelPackage.Save(password);
        }
        public void saveAs(string file)
        {
            saveAs(file, _excelPackage.Encryption.Password);
        }
        public void saveAs(Context ctx, PhpString file)
        {
            saveAs(file.ToString(ctx));
        }
        public void saveAs(Context ctx, PhpString file, PhpString password)
        {
            saveAs(file.ToString(ctx), password.ToString(ctx));
        }
        public void saveAs(string file, string password = null)
        {
            if (string.IsNullOrEmpty(password))
            {
                password = null;
            }
            checkWorkbookIsEmpty();
            _excelPackage.SaveAs(new FileInfo(file), password);
        }
        private void checkWorkbookIsEmpty()
        {
            // 空白工作簿将出错
            if (_excelPackage.Workbook.Worksheets.Count == 0)
            {
                _excelPackage.Workbook.Worksheets.Add("sheet1");
            }
        }

        public WorkBook workBook { get => _workBook; }
    }
}
