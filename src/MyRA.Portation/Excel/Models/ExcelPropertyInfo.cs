using System.Reflection;
using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Excel.Models
{
    /// <summary>
    ///     Class for information about to be parsed property.
    /// </summary>
    internal sealed class ExcelPropertyInfo
    {
        /// <summary>
        ///     Attribute of the property
        /// </summary>
        public ExcelPropertyAttribute Attribute { get; set; }

        /// <summary>
        ///     Column name of the to-be-parsed attribute (returns Property Name if ColumnName is not set)
        /// </summary>
        public string ColumnName => Attribute.ColumnName ?? Property.Name;

        /// <summary>
        ///     Property where attribute is applied to
        /// </summary>
        public PropertyInfo Property { get; set; }

        /// <summary>
        /// Gets or sets the column where data is exported to
        /// </summary>
        public int? Column { get; set; }

        public override string ToString()
        {
            return $"{ColumnName} ({Property.PropertyType}{(Column.HasValue ? $" at {Column}" : "")})";
        }
    }
}
