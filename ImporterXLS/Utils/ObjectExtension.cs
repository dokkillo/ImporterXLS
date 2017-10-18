using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ImporterXLS.Utils
{
    public static class ObjectExtensions
    {
        public static T ToObject<T>(this IDictionary<string, object> source)
            where T : class, new()
        {
            T someObject = new T();
            Type someObjectType = someObject.GetType();

            foreach (KeyValuePair<string, object> item in source)
            {
                var key = item.ToString();
                PropertyInfo propertyInfo = someObject.GetType().GetProperty(key);
                propertyInfo.SetValue(someObject, Convert.ChangeType(item.Value, propertyInfo.PropertyType), null);

            }

            return someObject;
        }
    }
}
