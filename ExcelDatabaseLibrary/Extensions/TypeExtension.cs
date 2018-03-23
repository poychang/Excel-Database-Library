using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelDatabaseLibrary.Extensions
{
    /// <summary>Type 類別的擴充方法</summary>
    public static class TypeExtension
    {
        /// <summary>取得此類別所指定的 Attribute 實體</summary>
        /// <typeparam name="T">指定的 Attribute</typeparam>
        /// <param name="instance">擴充方法對象</param>
        /// <returns>Attribute 實體</returns>
        public static T GetAttribute<T>(this Type instance) where T : Attribute
        {
            var attributeType = typeof(T);
            return (T)instance.GetCustomAttributes(attributeType, false).FirstOrDefault();
        }

        /// <summary>取得此類別所指定的 Attribute 實體屬性</summary>
        /// <typeparam name="T">指定的 Attribute</typeparam>
        /// <param name="instance">擴充方法對象</param>
        /// <param name="propertyName">屬性名稱</param>
        /// <returns>Attribute 實體</returns>
        public static T GetAttributeProperty<T>(this Type instance, string propertyName) where T : Attribute
        {
            var attributeType = typeof(T);
            var property = instance.GetProperty(propertyName);
            if (property == null) return default(T);
            return (T)property.GetCustomAttributes(attributeType, false).FirstOrDefault();
        }

        /// <summary>取得此類別所指定的 Attribute 實體屬性</summary>
        /// <typeparam name="T">指定的 Attribute</typeparam>
        /// <param name="instance">擴充方法對象</param>
        /// <returns>Attribute 實體</returns>
        public static IEnumerable<T> GetAttributeProperties<T>(this Type instance) where T : Attribute
        {
            var attributeType = typeof(T);
            var properties = instance.GetProperties();
            if (properties != null && properties.Length == 0) return default(IEnumerable<T>);
            return properties?.Select(property =>
                (T)property.GetCustomAttributes(attributeType, false).FirstOrDefault());
        }
    }
}
