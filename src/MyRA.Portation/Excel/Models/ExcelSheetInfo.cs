using System;
using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Excel.Models
{
    /// <summary>
    ///     Class for holding information about to be parsed sheet.
    /// </summary>
    internal abstract class ExcelSheetInfo
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelSheetInfo" /> class.
        /// </summary>
        /// <param name="attribute">The attribute.</param>
        /// <param name="type">The type.</param>
        protected ExcelSheetInfo(ExcelSheetAttribute attribute, Type type)
        {
            Attribute = attribute;
            Type = type;
        }

        /// <summary>
        ///     Attribute of the property
        /// </summary>
        public ExcelSheetAttribute Attribute { get; }

        /// <summary>
        ///     Column name of the to-be-parsed attribute (returns Property Name if ColumnName is not set)
        /// </summary>
        public virtual string SheetName => Attribute.SheetName;

        /// <summary>
        ///     Gets the type of model to be parsed.
        /// </summary>
        /// <value>
        ///     The type.
        /// </value>
        public Type Type { get; }

        /// <summary>
        ///     Get value to be exported.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>Value to be exported.</returns>
        public abstract object GetValue(object value);

        /// <summary>
        ///     Sets the value from import.
        /// </summary>
        /// <param name="model">The model.</param>
        /// <param name="value">The value.</param>
        public abstract void SetValue(ref object model, object value);

        public override string ToString()
        {
            return $"{SheetName}";
        }
    }
}