using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using MyRA.Portation.Excel.Attributes;
using MyRA.Portation.Excel.Models;
using MyRA.Portation.TypeConverters;

namespace MyRA.Portation.Excel
{
    internal static class ExcelReflection
    {
        /// <summary>
        ///     Gets the generic type from enumerable or null for other types.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        public static Type GetGenericTypeFromEnumerableOrDefault(Type type)
        {
            // string inherts form ienumerable<char>; and that is not what we want now
            return type == typeof(string)
                ? null
                : type.GetInterfaces().FirstOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>))?.GetGenericArguments()[0];
        }

        /// <summary>
        ///     Get all ParseProperty applied on public properties of type
        /// </summary>
        /// <param name="type">Type to parse all ParseProperty</param>
        /// <returns>IEnumerable of ParseProperty</returns>
        public static IList<ExcelPropertyInfo> GetParsingProperties(Type type)
        {
            var genericType = GetGenericTypeFromEnumerableOrDefault(type);
            var targetType = genericType ?? type;

            ExcelPropertyAttribute attribute = null;
            var properties = targetType.GetProperties()
                .Where(p => TryGetAttribute(p, out attribute))
                .Select(p => new ExcelPropertyInfo
                {
                    Property = p,
                    Attribute = attribute
                })
                .ToArray();

            for (var i = 0; i < properties.Length; ++i)
            {
                if (!properties[i].Attribute.Column.HasValue)
                    properties[i].Attribute.Column = i;
            }

            return properties;
        }

        /// <summary>
        ///     Gets the sheet information.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        public static ExcelSheetClassInfo GetSheetInfo(Type type)
        {
            var genericType = GetGenericTypeFromEnumerableOrDefault(type);
            var targetType = genericType ?? type;

            var attribute = targetType.GetCustomAttribute<ExcelSheetAttribute>();
            return attribute == null
                ? new ExcelSheetClassInfo(new ExcelSheetAttribute { SheetName = targetType.Name }, type)
                : new ExcelSheetClassInfo(attribute, type);
        }

        /// <summary>
        ///     Get all SheetProperty applied on public properties of type
        /// </summary>
        /// <param name="type">Type to parse</param>
        /// <returns>IEnumerable of SheetProperty</returns>
        public static IList<ExcelSheetPropertyInfo> GetSheetProperties(Type type)
        {
            var genericType = GetGenericTypeFromEnumerableOrDefault(type);
            var targetType = genericType ?? type;

            ExcelSheetAttribute attribute = null;
            var properties = targetType.GetProperties()
                .Where(p => TryGetAttribute(p, out attribute))
                .Select(p => new ExcelSheetPropertyInfo(attribute, p))
                .ToArray();

            return properties;
        }

        /// <summary>
        ///     Gets the value of ExcelPropertyInfo.
        /// </summary>
        /// <param name="parsingPropertyInfo">The parsing property information.</param>
        /// <param name="obj">The object.</param>
        /// <returns></returns>
        public static string GetValue(ExcelPropertyInfo parsingPropertyInfo, object obj)
        {
            var targetType = parsingPropertyInfo.Property.PropertyType;
            var converter = GetTypeConverter(parsingPropertyInfo, targetType);

            var value = parsingPropertyInfo.Property.GetValue(obj);
            return value is string str
                ? str
                : converter.ConvertToInvariantString(value);
        }

        /// <summary>
        ///     Converts value to the correct type and set the value on public property of obj
        /// </summary>
        /// <param name="parsingPropertyInfo">The property to be set</param>
        /// <param name="obj">The object on which the property should be set</param>
        /// <param name="value">Value which should be converted and set</param>
        public static void SetConvertedValue(ExcelPropertyInfo parsingPropertyInfo, object obj, string value)
        {
            var targetType = parsingPropertyInfo.Property.PropertyType;
            var nullableTargetType = Nullable.GetUnderlyingType(targetType);

            // nullable type
            if (nullableTargetType != null)
            {
                // if string is null, set it as null
                if (string.IsNullOrEmpty(value))
                {
                    parsingPropertyInfo.Property.SetValue(obj, null);
                    return;
                }

                targetType = nullableTargetType;
            }

            // find the converter
            var converter = GetTypeConverter(parsingPropertyInfo, targetType);

            // convert value
            var convertedValue = converter.ConvertFromInvariantString(value);
            parsingPropertyInfo.Property.SetValue(obj, convertedValue);
        }

        private static TypeConverter GetTypeConverter(ExcelPropertyInfo parsingPropertyInfo, Type targetType)
        {
            var converterType = parsingPropertyInfo.Attribute.Converter;
            var converter = converterType?.IsAbstract == false && converterType.IsSubclassOf(typeof(TypeConverter))
                ? (TypeConverter)Activator.CreateInstance(converterType)
                : TypeDescriptor.GetConverter(targetType);

            if (converter is IFormattableConverter formatableConverter && !string.IsNullOrEmpty(parsingPropertyInfo.Attribute.ConverterFormat))
                formatableConverter.ConvertFormat = parsingPropertyInfo.Attribute.ConverterFormat;

            return converter;
        }

        private static bool TryGetAttribute<T>(ICustomAttributeProvider memberInfo, out T customAttribute) where T : Attribute
        { // Try to get attribute of T from the memberInfo (Properties, fields etc.)
            var attributes = memberInfo.GetCustomAttributes(typeof(T), false).FirstOrDefault();
            if (attributes == null)
            {
                customAttribute = null;
                return false;
            }

            customAttribute = (T)attributes;
            return true;
        }
    }
}
