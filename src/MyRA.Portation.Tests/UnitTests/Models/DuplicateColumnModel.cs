using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Tests.UnitTests.Models
{
    internal sealed class DuplicateColumnModel
    {
        [ExcelProperty(1)]
        public int Id { get; set; }

        [ExcelProperty(1)]
        public string Firstname { get; set; }
    }
}
