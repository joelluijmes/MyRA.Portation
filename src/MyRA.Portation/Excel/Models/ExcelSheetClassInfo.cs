using System;
using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Excel.Models
{
    /// <summary>
    ///     Class for holding information of to be parsed class model.
    /// </summary>
    /// <seealso cref="ExcelSheetInfo" />
    internal sealed class ExcelSheetClassInfo : ExcelSheetInfo
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelSheetClassInfo" /> class.
        /// </summary>
        /// <param name="attribute">The attribute.</param>
        /// <param name="type">The type.</param>
        public ExcelSheetClassInfo(ExcelSheetAttribute attribute, Type type) : base(attribute, type) { }

        /// <summary>
        ///     Get value to be exported.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>
        ///     value
        /// </returns>
        public override object GetValue(object value)
        {
            return value;
        }

        /// <summary>
        ///     Sets the value from import.
        /// </summary>
        /// <param name="model">The model will be overwritten by value.</param>
        /// <param name="value">The value.</param>
        public override void SetValue(ref object model, object value)
        {
            model = value;
        }
    }
}
