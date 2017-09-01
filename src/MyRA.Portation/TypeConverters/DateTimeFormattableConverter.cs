using System;
using System.ComponentModel;
using System.Globalization;
using MyRA.Portation.Exceptions;

namespace MyRA.Portation.TypeConverters
{
    /// <summary>
    ///     Formattable DateTime converter, the Format of DateTime is injected at runtime during parsing.
    /// </summary>
    /// <seealso cref="System.ComponentModel.TypeConverter" />
    /// <seealso cref="IFormattableConverter" />
    public sealed class DateTimeFormattableConverter : TypeConverter, IFormattableConverter
    {
        public DateTimeFormattableConverter() { }

        public DateTimeFormattableConverter(string convertFormat)
        {
            ConvertFormat = convertFormat;
        }

        public string ConvertFormat { get; set; }

        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return sourceType == typeof(string) || base.CanConvertFrom(context, sourceType);
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            return destinationType == typeof(DateTime) || base.CanConvertTo(context, destinationType);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            if (string.IsNullOrEmpty(ConvertFormat))
                throw new ParserException($"{nameof(DateTimeFormattableConverter)} implements {nameof(IFormattableConverter)}, however {nameof(ConvertFormat)} is null or empty");

            switch (value)
            {
            case string strValue:
                return DateTime.ParseExact(strValue, ConvertFormat, culture);

            default:
                return base.ConvertFrom(context, culture, value);
            }
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            if (string.IsNullOrEmpty(ConvertFormat))
                throw new ParserException($"{nameof(DateTimeFormattableConverter)} implements {nameof(IFormattableConverter)}, however {nameof(ConvertFormat)} is null or empty");

            switch (value)
            {
            case DateTime date:
                return date.ToString(ConvertFormat);

            default:
                return base.ConvertFrom(context, culture, value);
            }
        }
    }
}
