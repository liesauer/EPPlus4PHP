using System;

namespace nulastudio.Document.EPPlus4PHP
{
    public enum DataType
    {
        // 整数
        Integer,
        // 浮点数
        Decimal,
        // 字符串
        String,
        // 布尔值
        Boolean,
        // 日期类型（无法字面赋值）
        Date,
        // 时间类型（无法字面赋值）
        Time,
        // 可迭代类型（无法字面赋值）
        Enumerable,
        // 数组（无法字面赋值）
        LookupArray,
        // 单元格
        ExcelAddress,
        // 错误（无法字面赋值）
        ExcelError,
        // 空白
        Empty,
        // 未知（无法字面赋值）
        Unknown,
    }
}