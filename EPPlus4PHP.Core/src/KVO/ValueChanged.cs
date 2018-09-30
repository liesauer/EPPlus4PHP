using System;

namespace nulastudio.KVO
{
    public class ValueChanged
    {
        public event EventHandler<ValueChangedEventArgs> OnValueChanged;
        public bool hasEvent => OnValueChanged != null;

        public virtual void TriggerValueChanged(ValueChangedEventArgs e)
        {
            if (OnValueChanged != null)
            {
                OnValueChanged(this, e);
            }
        }
    }
}