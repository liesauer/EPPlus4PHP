using System;
using Pchp.Core;
using Pchp.Library;
using nulastudio.KVO;

namespace nulastudio.Document.EPPlus4PHP.Style
{
    public class NumberFormat : ValueChanged
    {
        private string _format;

        public NumberFormat() : this(@"")
        {
        }
        public NumberFormat(string format)
        {
            _format = format;
        }

        public string format
        {
            get => _format;
            set
            {
                string oldValue = _format;
                _format = value;
                TriggerValueChanged(new ValueChangedEventArgs("format", oldValue, _format));
            }
        }
        // will not trigger valuechanged
        public void setFormat(string foramt)
        {
            _format = format;
        }
    }
}