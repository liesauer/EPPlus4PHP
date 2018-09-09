using System;

namespace nulastudio.KVO
{
    public class ValueChanged
    {
        public event EventHandler<ValueChangedEventArgs> OnValueChanged;

        public virtual void TriggerValueChanged(ValueChangedEventArgs e)
        {
            if (OnValueChanged != null)
            {
                OnValueChanged(this, e);
            }
        }
    }
}