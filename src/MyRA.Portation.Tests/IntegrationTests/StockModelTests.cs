using System.IO;
using MyRA.Portation.Excel.Exporters;
using MyRA.Portation.Excel.Importers;
using MyRA.Portation.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace MyRA.Portation.Tests.IntegrationTests
{
    public sealed class StockModelTests
    {
        [Fact]
        public void TestObjectStock_Export()
        {
            const string item1 = "Water bottle";
            const string item2 = "Coffee cup";
            const string item3 = "iPhone";

            var items = new[]
            {
                new ItemModel {Name = item1},
                new ItemModel {Name = item2},
                new ItemModel {Name = item3},
            };

            const string warehouse = "Topicus";
            var stock = new StockModel
            {
                Items = items,
                Warehouse = warehouse
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(stock);

                // NOTE - can't really mock the excel export validation.. just use EPPlus if we can read the data 
                var package = new ExcelPackage(stream);

                // two sheets, one for items one for person
                Assert.Equal(2, package.Workbook.Worksheets.Count);

                // NOTE - EPPlus starts at 1
                var infoWorksheet = package.Workbook.Worksheets[1];
                Assert.Equal(StockModel.SHEET_NAME, infoWorksheet.Name);
                Assert.Equal(nameof(StockModel.Warehouse), infoWorksheet.Cells[1, 1].Text);
                Assert.Equal(warehouse, infoWorksheet.Cells[1, 2].Text);
                
                var itemsWorksheet = package.Workbook.Worksheets[2];
                Assert.Equal(nameof(OrderModel.Items), itemsWorksheet.Name);
                Assert.Equal(nameof(ItemModel.Name), itemsWorksheet.Cells[1, 1].Text);
                Assert.Equal(item1, itemsWorksheet.Cells[2, 1].Text);
                Assert.Equal(item2, itemsWorksheet.Cells[3, 1].Text);
                Assert.Equal(item3, itemsWorksheet.Cells[4, 1].Text);
            }
        }

        [Fact]
        public void TestObjectStock_Import()
        {
            const string item1 = "Water bottle";
            const string item2 = "Coffee cup";
            const string item3 = "iPhone";

            var items = new[]
            {
                new ItemModel {Name = item1},
                new ItemModel {Name = item2},
                new ItemModel {Name = item3},
            };

            const string warehouse = "Topicus";
            var stock = new StockModel
            {
                Items = items,
                Warehouse = warehouse
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(stock);

                var importer = new ExcelModelImporter(stream);
                var outputOrder = importer.ImportModel<StockModel>();

                Assert.Equal(stock, outputOrder);
            }
        }
    }
}
