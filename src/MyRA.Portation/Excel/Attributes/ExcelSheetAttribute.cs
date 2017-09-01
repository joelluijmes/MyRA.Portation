using System;

namespace MyRA.Portation.Excel.Attributes
{
    /// <summary>
    ///     Attribute for indicating sheet name, it also can be used to parse more complex objects.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Class)]
    public sealed class ExcelSheetAttribute : Attribute
    {
        /// <summary>
        ///     Name for parsing the correct excel sheet
        /// </summary>
        public string SheetName { get; set; }

        public override string ToString()
        {
            return $"{SheetName}";
        }
    }
}