namespace nulastudio.KVO
{
    public class ValueChangedEventArgs
    {
        private string _propertyName;
        public string PropertyName { get; }
        private object _oldValue;
        public object OldValue { get; }
        private object _newValue;
        public object NewValue { get; }

        public ValueChangedEventArgs(string propertyName, object oldValue, object newValue)
        {
               _propertyName = propertyName;
               _oldValue = oldValue;
               _newValue = newValue;
        }
    }
}