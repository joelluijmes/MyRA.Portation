using System;

namespace MyRA.Portation.Excel.Attributes
{
    /// <summary>
    ///     Attribute to indicate that this property should be parsed
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelPropertyAttribute : Attribute
    {
        /// <summary>
        ///     Static column position
        /// </summary>
        public int? Column { get; set; }

        /// <summary>
        ///     <para>Cell name to check, if not set uses property name</para>
        ///     <para>Note: this is case sensitive</para>
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        ///     <para>Converter to convert the value to a type.</para>
        ///     <para>The converter must implement IParseConverter</para>
        ///     <para>Note: this property is a type i.e typeof(DateParseConverter)</para>
        /// </summary>
        public Type Converter { get; set; }

        /// <summary>
        ///     Format string injected if Converter implements IFormatableConverter
        /// </summary>
        public string ConverterFormat { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelPropertyAttribute"/> class.
        /// </summary>
        public ExcelPropertyAttribute() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelPropertyAttribute"/> class.
        /// </summary>
        /// <param name="column">The column.</param>
        public ExcelPropertyAttribute(int column)
        {
            Column = column;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelPropertyAttribute"/> class.
        /// </summary>
        /// <param name="columnName">Name of the column.</param>
        public ExcelPropertyAttribute(string columnName)
        {
            ColumnName = columnName;
        }

        public override string ToString()
        {
            return $"{ColumnName}";
        }
    }
}
