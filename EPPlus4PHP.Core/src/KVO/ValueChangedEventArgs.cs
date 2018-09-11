namespace nulastudio.KVO
{
    public class ValueChangedEventArgs
    {
        private string _propertyName;
        public string PropertyName { get => _propertyName; }
        private object _oldValue;
        public object OldValue { get => _oldValue; }
        private object _newValue;
        public object NewValue { get => _newValue; }

        public ValueChangedEventArgs(string propertyName, object oldValue, object newValue)
        {
               _propertyName = propertyName;
               _oldValue = oldValue;
               _newValue = newValue;
        }
    }
}