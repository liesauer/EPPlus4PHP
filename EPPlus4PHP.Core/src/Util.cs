// using System;
// using System.Collections.Generic;
// using Pchp.Core;
// using Pchp.Library;
// using OfficeOpenXml;
// using OfficeOpenXml.FormulaParsing;
// using OfficeOpenXml.FormulaParsing.Excel.Functions;
// using OfficeOpenXml.FormulaParsing.ExpressionGraph;
// using ExcelDataType = OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType;
// using OfficeOpenXml.FormulaParsing.Exceptions;

// namespace nulastudio.Document.EPPlus4PHP
// {
//     public class Util
//     {
//         public static void createResult(Context ctx, object value, DataType dataType)
//         {
//             object val;
//             ExcelDataType dt = ExcelDataType.Unknown;
//             // int or long
//             if (dataType == DataType.Integer)
//             {
//                 dt = ExcelDataType.Integer;
//                 if (value is int || value is long)
//                 {
//                     val = value;
//                 }
//                 else if (value is float || value is double)
//                 {
//                     val = (int)(double)value;
//                 }
//                 else if (value is string)
//                 {
//                     if (long.TryParse(value as string, out var lnum))
//                     {
//                         val = lnum;
//                     }
//                     else if (int.TryParse(value as string, out var inum))
//                     {
//                         val = inum;
//                     }
//                 }
//                 else if (value is PhpString)
//                 {
//                     string str = ((PhpString)value).ToString(ctx);
//                     if (long.TryParse(str, out var lnum))
//                     {
//                         val = lnum;
//                     }
//                     else if (int.TryParse(str, out var inum))
//                     {
//                         val = inum;
//                     }
//                 }
//             }
//         }
//     }
// }