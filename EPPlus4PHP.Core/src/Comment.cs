using System;
using System.Collections.Generic;
using System.Linq;
using Pchp.Core;
using Pchp.Library;
using OfficeOpenXml;
using nulastudio.KVO;

namespace nulastudio.Document.EPPlus4PHP
{
    public class Comment : ValueChanged
    {
        private string _text;
        private string _author;
        public Comment(string text) : this(text, "Author")
        {}
        public Comment(string text, string author)
        {
            _text = text;
            _author = author;
        }

        public string text {
            get => _text;
            set
            {
                string oldVal = _text;
                _text = value;
                TriggerValueChanged(new ValueChangedEventArgs("text", oldVal, _text));
            }
        }

        public string author
        {
            get => _author;
            set
            {
                string oldVal = _author;
                _author = value;
                TriggerValueChanged(new ValueChangedEventArgs("author", oldVal, _author));
            }
        }

        public static implicit operator Comment(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }
            return new Comment(text);
        }
    }
}