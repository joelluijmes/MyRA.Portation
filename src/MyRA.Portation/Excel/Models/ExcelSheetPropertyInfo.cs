using System.Reflection;
using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Excel.Models
{
    /// <summary>
    ///     Class for holding information about property to be parsed.
    /// </summary>
    /// <seealso cref="ExcelSheetInfo" />
    internal sealed class ExcelSheetPropertyInfo : ExcelSheetInfo
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelSheetPropertyInfo" /> class.
        /// </summary>
        /// <param name="attribute">The attribute.</param>
        /// <param name="property">The property.</param>
        public ExcelSheetPropertyInfo(ExcelSheetAttribute attribute, PropertyInfo property) : base(attribute,
            property.PropertyType)
        {
            Property = property;
        }

        /// <summary>
        ///     Property where attribute is applied to
        /// </summary>
        public PropertyInfo Property { get; set; }

        /// <summary>
        ///     Sheet name of the to-be-parsed attribute (returns Property Name if Attribute.SheetName is not set)
        /// </summary>
        public override string SheetName => Attribute.SheetName ?? Property.Name;

        /// <summary>
        ///     Get value to be exported.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>
        ///     Property.GetValue(value);
        /// </returns>
        public override object GetValue(object value)
        {
            return Property.GetValue(value);
        }

        /// <summary>
        ///     Sets the value from import.
        /// </summary>
        /// <param name="model">The model.</param>
        /// <param name="value">The value.</param>
        public override void SetValue(ref object model, object value)
        {
            Property.SetValue(model, value);
        }

        public override string ToString()
        {
            return $"{SheetName} ({Property.PropertyType})";
        }
    }
}