using MyRA.Portation.Excel.Attributes;

namespace MyRA.Portation.Tests.UnitTests.Models
{
    internal sealed class InvalidColumnModel
    {
        [ExcelProperty(0)]
        public int Id { get; set; }
    }
}
