using System;
using System.Collections.Generic;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using ExcelDataType = OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace nulastudio.Document.EPPlus4PHP
{
    public class Function : ExcelFunction
    {
        private Context _ctx;
        private ExcelPackage _package;
        private Closure _callback;
        private Function()
        {}
        public Function(Context ctx, ExcelPackage package, Closure callback)
        {
            _ctx = ctx;
            _package = package;
            _callback = callback;
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            // 参数转换
            // 1. 转换EPPlus内部的类型
            // 2. 转换EPPlus4PHP的类型（PhpValue类型，不用处理）
            PhpArray args = PhpArray.NewEmpty();
            foreach (FunctionArgument arg in arguments)
            {
                PhpValue val = PhpValue.Null;

                if (arg.Value is PhpValue v)
                {
                    val = v;
                }
                else /* EPPlus内部类型 */
                {
                    if (arg.IsExcelRange && arg.Value is EpplusExcelDataProvider.RangeInfo rangeInfo)
                    {
                        string rangeAddress = rangeInfo.Address.Address;
                        Range range = new Range(rangeInfo.Worksheet.Cells[rangeAddress], _package.is1base);
                        val = PhpValue.FromClr(range);
                    }
                    else
                    {
                        val = PhpValue.FromClr(arg.Value);
                    }
                }
                args.AddValue(val);
            }

            // 上下文
            PhpArray contextArr = PhpArray.NewEmpty();
            // context.Scopes.Current.Address对于区域地址有问题
            // 暂时不用
            // contextArr.Add(context.Scopes.Current.Address);

            // 获取结果
            PhpValue ret = _callback.__invoke(PhpValue.Create(args), PhpValue.Create(contextArr));

            // 结果可能存在两种类型
            // 1. 可直接转换的类型或者可直接字面表达的值
            // 2. Result类型

            // 将结果转换为Result类型

            // 将Result转换为EPPlus的Result类型

            // object val = null;
            // ExcelDataType dt = ExcelDataType.Empty;
            return CreateResult(ret, ExcelDataType.Unknown);
        }
    }
}