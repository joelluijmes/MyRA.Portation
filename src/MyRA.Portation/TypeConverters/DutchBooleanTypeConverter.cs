using System;
using System.ComponentModel;
using System.Globalization;

namespace MyRA.Portation.TypeConverters
{
    /// <summary>
    ///     Dutch boolean TypeConverter
    /// </summary>
    /// <seealso cref="System.ComponentModel.TypeConverter" />
    public sealed class DutchBooleanTypeConverter : TypeConverter
    {
        /// <summary>
        ///     Returns whether this converter can convert an object of the given type to the type of this converter, using the
        ///     specified context.
        /// </summary>
        /// <param name="context">
        ///     An <see cref="T:System.ComponentModel.ITypeDescriptorContext"></see> that provides a format
        ///     context.
        /// </param>
        /// <param name="sourceType">A <see cref="T:System.Type"></see> that represents the type you want to convert from.</param>
        /// <returns>
        ///     true if this converter can perform the conversion; otherwise, false.
        /// </returns>
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return sourceType == typeof(string) || base.CanConvertFrom(context, sourceType);
        }

        /// <summary>
        ///     Returns whether this converter can convert the object to the specified type, using the specified context.
        /// </summary>
        /// <param name="context">
        ///     An <see cref="T:System.ComponentModel.ITypeDescriptorContext"></see> that provides a format
        ///     context.
        /// </param>
        /// <param name="destinationType">A <see cref="T:System.Type"></see> that represents the type you want to convert to.</param>
        /// <returns>
        ///     true if this converter can perform the conversion; otherwise, false.
        /// </returns>
        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            return destinationType == typeof(bool) || base.CanConvertTo(context, destinationType);
        }

        /// <summary>
        ///     Converts the given object to the type of this converter, using the specified context and culture information.
        /// </summary>
        /// <param name="context">
        ///     An <see cref="T:System.ComponentModel.ITypeDescriptorContext"></see> that provides a format
        ///     context.
        /// </param>
        /// <param name="culture">The <see cref="T:System.Globalization.CultureInfo"></see> to use as the current culture.</param>
        /// <param name="value">The <see cref="T:System.Object"></see> to convert.</param>
        /// <returns>
        ///     An <see cref="T:System.Object"></see> that represents the converted value.
        /// </returns>
        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            switch (value)
            {
                case string strValue:
                    switch (strValue.ToLower())
                    {
                        case "ja":
                        case "j":
                        case "1":
                        case "y":
                        case "yes":
                        case "true":
                        case "t":
                            return true;

                        case "nee":
                        case "n":
                        case "0":
                        case "no":
                        case "false":
                        case "f":
                            return false;

                        default:
                            return bool.Parse(strValue);
                    }

                default:
                    return base.ConvertFrom(context, culture, value);
            }
        }

        /// <summary>
        ///     Converts the given value object to the specified type, using the specified context and culture information.
        /// </summary>
        /// <param name="context">
        ///     An <see cref="T:System.ComponentModel.ITypeDescriptorContext"></see> that provides a format
        ///     context.
        /// </param>
        /// <param name="culture">
        ///     A <see cref="T:System.Globalization.CultureInfo"></see>. If null is passed, the current culture
        ///     is assumed.
        /// </param>
        /// <param name="value">The <see cref="T:System.Object"></see> to convert.</param>
        /// <param name="destinationType">The <see cref="T:System.Type"></see> to convert the value parameter to.</param>
        /// <returns>
        ///     An <see cref="T:System.Object"></see> that represents the converted value.
        /// </returns>
        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            switch (value)
            {
                case bool boolValue:
                    return boolValue ? "ja" : "nee";

                default:
                    return base.ConvertFrom(context, culture, value);
            }
        }
    }
}
