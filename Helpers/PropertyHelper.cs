using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace ExcelProvider.Helpers
{
    public static class PropertyHelper
    {
        private static readonly List<Type> NumericTypes = new List<Type>
        {
            typeof(byte),
            typeof(short),
            typeof(int),
            typeof(decimal),
            typeof(float),
            typeof(double)
        };

        private static bool IsNumericType(Type type)
        {
            return NumericTypes.Contains(type) || NumericTypes.Contains(Nullable.GetUnderlyingType(type));
        }

        public static string GetPropertyName(this PropertyInfo property)
        {
            return (property.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as DisplayNameAttribute)?.DisplayName ?? property.Name;
        }

        public static object GetPropertyValue(this object model, string propertyName)
        {
            return model.GetType().GetProperties().FirstOrDefault(p => p.GetPropertyName() == propertyName)?.GetValue(model);
        }

        public static IEnumerable<string> GetPropertiesNames(this Type type, bool onlyMapped = true)
        {
            return type.GetProperties()
                .Where(p => !onlyMapped || p.GetCustomAttributes(typeof(NotMappedAttribute), false).FirstOrDefault() == null)
                .Select(p => p.GetPropertyName());
        }

        public static void SetPropertyValue<TModel>(this PropertyInfo property, TModel model, string value)
        {
            if (IsNumericType(property.PropertyType))
            {
                value = new string((value as string).ToCharArray().Where(c => !char.IsWhiteSpace(c)).ToArray()).Replace(",", ".");
            }

            property.SetValue(model, Convert.ChangeType(value, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType, new CultureInfo("en-US")));
        }
    }
}
