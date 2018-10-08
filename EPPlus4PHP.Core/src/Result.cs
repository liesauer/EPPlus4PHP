using System;
using System.Collections.Generic;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace nulastudio.Document.EPPlus4PHP
{
    public class Result
    {
        private object _value = null;
        private DataType _dataType = DataType.Empty;

        public object value { get => _value; }
        public DataType dataType { get => _dataType; }

        private Result(object value, DataType dataType)
        {
            _value = value;
            _dataType = dataType;
        }

        public static Result create(Context ctx, object value, int dataType)
        {
            object val = null;
            DataType dt = DataType.Unknown;

            try
            {
                dt = (DataType)dataType;
                switch ((DataType)dataType)
                {
                    case DataType.Integer:
                        dt = ExcelDataType.Integer;
                        if (value is int || value is long)
                        {
                            val = value;
                        }
                        else if (value is float || value is double)
                        {
                            val = (int)(double)value;
                        }
                        else if (value is string)
                        {
                            if (long.TryParse(value as string, out var lnum))
                            {
                                val = lnum;
                            }
                            else if (int.TryParse(value as string, out var inum))
                            {
                                val = inum;
                            }
                        }
                        else if (value is PhpString)
                        {
                            string str = ((PhpString)value).ToString(ctx);
                            if (long.TryParse(str, out var lnum))
                            {
                                val = lnum;
                            }
                            else if (int.TryParse(str, out var inum))
                            {
                                val = inum;
                            }
                        }
                        break;
                    case DataType.Decimal:
                        break;
                    case DataType.String:
                        break;
                    case DataType.Boolean:
                        break;
                    case DataType.Date:
                        break;
                    case DataType.Time:
                        break;
                    case DataType.Enumerable:
                        break;
                    case DataType.LookupArray:
                        break;
                    case DataType.ExcelAddress:
                        break;
                    case DataType.ExcelError:
                        break;
                    case DataType.Empty:
                        val = null;
                        dt = DataType.Empty;
                        break;
                    case DataType.Unknown:
                    default:
                        val = null;
                        dt = DataType.Unknown;
                        break;
                }
            }
            catch (ExcelErrorValueException e)
            {
                val = new ErrorValue((ErrorValueType)(int)e.ErrorValue.Type);
                dt = DataType.ExcelError;
            }
            catch (Exception)
            {
                val = new ErrorValue(ErrorValueType.Value);
                dt = DataType.ExcelError;
            }
            return create(val, dt);
        }
        public static Result create(object value, DataType dataType)
        {
            return new Result(value, dataType);
        }
    }
}