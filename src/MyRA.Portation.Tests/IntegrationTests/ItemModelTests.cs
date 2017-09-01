using System.Collections.Generic;
using System.IO;
using MyRA.Portation.Excel;
using MyRA.Portation.Excel.Exporters;
using MyRA.Portation.Excel.Importers;
using MyRA.Portation.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace MyRA.Portation.Tests.IntegrationTests
{
    public sealed class ItemModelTests
    {
        [Fact]
        public void TestObjectItems_Export()
        {
            const string item1 = "Water bottle";
            const string item2 = "Coffee cup";
            const string item3 = "iPhone";

            var items = new[]
            {
                new ItemModel {Name = item1},
                new ItemModel {Name = item2},
                new ItemModel {Name = item3}
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(items);

                // NOTE - can't really mock the excel export validation.. just use EPPlus if we can read the data 
                var package = new ExcelPackage(stream);

                // NOTE - EPPlus starts at 1
                Assert.Equal(1, package.Workbook.Worksheets.Count);
                var itemsWorksheet = package.Workbook.Worksheets[1];

                Assert.Equal(nameof(ItemModel), itemsWorksheet.Name);
                Assert.Equal(nameof(ItemModel.Name), itemsWorksheet.Cells[1, 1].Text);
                Assert.Equal(item1, itemsWorksheet.Cells[2, 1].Text);
                Assert.Equal(item2, itemsWorksheet.Cells[3, 1].Text);
                Assert.Equal(item3, itemsWorksheet.Cells[4, 1].Text);
            }
        }

        [Fact]
        public void TestObjectItems_Import()
        {
            const string item1 = "Water bottle";
            const string item2 = "Coffee cup";
            const string item3 = "iPhone";

            var items = new[]
            {
                new ItemModel {Name = item1},
                new ItemModel {Name = item2},
                new ItemModel {Name = item3}
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(items);

                var importer = new ExcelModelImporter(stream);
                var otuputItems = importer.ImportModel<List<ItemModel>>();

                Assert.Equal(items, otuputItems);
            }
        }

        [Fact]
        public void TestObjectItems_ParsingProperties()
        {
            var targetType = typeof(List<ItemModel>);
            var parsingProperties = ExcelReflection.GetParsingProperties(targetType);

            Assert.Equal(1, parsingProperties.Count);
            Assert.Equal(nameof(ItemModel.Name), parsingProperties[0].ColumnName);
        }
    }
}