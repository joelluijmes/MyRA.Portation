using System.IO;
using MyRA.Portation.Excel;
using MyRA.Portation.Excel.Exporters;
using MyRA.Portation.Excel.Importers;
using MyRA.Portation.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace MyRA.Portation.Tests.IntegrationTests
{
    public sealed class OrderModelTests
    {
        [Fact]
        public void TestObjectOrder_GetSheetName()
        {
            var targetType = typeof(OrderModel);
            var orderSheetInfo = ExcelReflection.GetSheetInfo(targetType);

            Assert.Equal(OrderModel.SHEET_NAME, orderSheetInfo.SheetName);
        }

        [Fact]
        public void TestObjectOrder_Export()
        {
            const string person1Firstname = "Anne";
            const string person1Lastname = "Jenzen";

            const string item1 = "Water bottle";
            const string item2 = "Coffee cup";
            const string item3 = "iPhone";

            var items = new[]
            {
                new ItemModel {Name = item1},
                new ItemModel {Name = item2},
                new ItemModel {Name = item3},
            };

            var person = new PersonModel { Id = 1, Firstname = person1Firstname, Lastname = person1Lastname };

            var order = new OrderModel
            {
                Items = items,
                Person = person
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(order);

                // NOTE - can't really mock the excel export validation.. just use EPPlus if we can read the data 
                var package = new ExcelPackage(stream);

                // two sheets, one for items one for person
                Assert.Equal(2, package.Workbook.Worksheets.Count);

                // NOTE - EPPlus starts at 1
                var personWorksheet = package.Workbook.Worksheets[1];
                Assert.Equal(OrderModel.PERSON_SHEET_NAME, personWorksheet.Name);

                // NOTE - single object get exported as KEY : VALU instea of rows
                Assert.Equal(nameof(PersonModel.Id), personWorksheet.Cells[1, 1].Text);
                Assert.Equal(nameof(PersonModel.Firstname), personWorksheet.Cells[2, 1].Text);
                Assert.Equal(nameof(PersonModel.Lastname), personWorksheet.Cells[3, 1].Text);

                Assert.Equal(person1Firstname, personWorksheet.Cells[2, 2].Text);
                Assert.Equal(person1Lastname, personWorksheet.Cells[3, 2].Text);

                var itemsWorksheet = package.Workbook.Worksheets[2];
                Assert.Equal(nameof(OrderModel.Items), itemsWorksheet.Name);
                Assert.Equal(nameof(ItemModel.Name), itemsWorksheet.Cells[1, 1].Text);
                Assert.Equal(item1, itemsWorksheet.Cells[2, 1].Text);
                Assert.Equal(item2, itemsWorksheet.Cells[3, 1].Text);
                Assert.Equal(item3, itemsWorksheet.Cells[4, 1].Text);
            }
        }

        [Fact]
        public void TestObjectOrder_Import()
        {
            const string person1Firstname = "Anne";
            const string person1Lastname = "Jenzen";

            const string item1 = "Water bottle";
            const string item2 = "Coffee cup";
            const string item3 = "iPhone";

            var items = new[]
            {
                new ItemModel {Name = item1},
                new ItemModel {Name = item2},
                new ItemModel {Name = item3},
            };

            var person = new PersonModel { Id = 1, Firstname = person1Firstname, Lastname = person1Lastname };

            var order = new OrderModel
            {
                Items = items,
                Person = person
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(order);

                var importer = new ExcelModelImporter(stream);
                var outputOrder = importer.ImportModel<OrderModel>();

                Assert.Equal(order, outputOrder);
            }
        }
    }
}
