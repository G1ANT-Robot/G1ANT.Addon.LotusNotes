namespace G1ANT.Addon.LotusNotes.Models
{
    public class PropertyModel
    {
        public string Name { get; set; }
        public object Value { get; set; }

        public PropertyModel(string name, object value)
        {
            Name = name;
            Value = value;
        }
    }
}
