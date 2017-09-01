using MyRA.Portation.Excel;
using MyRA.Portation.Exceptions;
using MyRA.Portation.Tests.Models;
using MyRA.Portation.Tests.UnitTests.Models;
using Xunit;

namespace MyRA.Portation.Tests.UnitTests
{
    public sealed class ExcelPropertyTests
    {
        [Fact]
        public void ColumnIndexIsSequential()
        {
            var parsingProperties = ExcelReflection.GetParsingProperties(typeof(PersonModel));
            ExcelReflection.AutoAssignColumns(parsingProperties);

            for (var i = 0; i < parsingProperties.Count; ++i)
                Assert.Equal(i + 1, parsingProperties[i].Column);
        }

        [Fact]
        public void DuplicateColumn_ThrowsException()
        {
            Assert.Throws<ParserException>(
                () => { ExcelReflection.GetParsingProperties(typeof(DuplicateColumnModel)); });
        }

        [Fact]
        public void InvalidColumn_ThrowsException()
        {
            Assert.Throws<ParserException>(() => { ExcelReflection.GetParsingProperties(typeof(InvalidColumnModel)); });
        }
    }
}