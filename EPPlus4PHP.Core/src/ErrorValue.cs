using System;
using System.Collections;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;

namespace nulastudio.Document.EPPlus4PHP
{
    public class ErrorValue
    {
        public const string DIV0 = "#DIV/0!";
        public const string NA = "#N/A";
        public const string NAME = "#NAME?";
        public const string NULL = "#NULL!";
        public const string NUM = "#NUM!";
        public const string REF = "#REF!";
        public const string VALUE = "#VALUE!";

        public ErrorValue(ErrorValueType errorValueType)
        {
            _errorValueType = errorValueType;
        }

        private ErrorValueType _errorValueType;
        public int errorValueType { get => (int)_errorValueType; }
    }
}