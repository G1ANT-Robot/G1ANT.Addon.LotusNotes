using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace G1ANT.Addon.LotusNotes.Models
{
    public class PropertiesContainer
    {
        public virtual IReadOnlyCollection<string> GetPropertyNames() => GetPropertyTypes().Select(t => t.Name).ToList();

        public T TryGet<T>(Func<T> f)
        {
            try { return f(); }
            catch { return default(T); }
        }


        private IEnumerable<PropertyInfo> GetPropertyTypes()
        {
            return GetType()
                .GetProperties()
                .Where(t => !t.PropertyType.Name.StartsWith("Lazy") && !t.PropertyType.Name.StartsWith("Func"));
        }

        public virtual object GetPropertyValue(string name)
        {
            var properties = GetAllProperties();
            return properties.FirstOrDefault(p => p.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase)).Value
                ?? throw new IndexOutOfRangeException($"Property '{name}' is not supported");
        }

        public virtual void SetPropertyValue(string name, object value)
        {
            var propertyName = GetPropertyNames().FirstOrDefault(pn => pn.Equals(name, StringComparison.CurrentCultureIgnoreCase));
            if (propertyName == null)
                throw new IndexOutOfRangeException($"Property '{name}' is not supported");

            try
            {
                var property = GetType().GetProperty(propertyName);
                property.SetValue(this, value);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Property '{name}' is readonly", ex);
            }

        }

        public virtual IReadOnlyList<PropertyModel> GetAllProperties()
        {
            return GetPropertyTypes()
                .Select(t => new PropertyModel(t.Name, t.GetValue(this)))
                .ToList();
        }

        public string GetPropertyValue(object value)
        {
            if (value is PropertiesContainer propertiesContainer)
                return $"{propertiesContainer.ToPropertyString()}";
            if (!(value is string) && value is IEnumerable enumerable)
                return $"Array[{string.Join(", ", enumerable.Cast<object>().Select(o => GetPropertyValue(o)))}]";
            else
                return value?.ToString();
        }

        public virtual string ToPropertyString()
        {
            return string.Join($", {Environment.NewLine}", GetAllProperties().Select(p => $"{p.Name}: {GetPropertyValue(p.Value)}"));
        }
    }
}
