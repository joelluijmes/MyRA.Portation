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
    public sealed class PersonModelTests
    {
        [Fact]
        public void TestObjectPerson_Export()
        {
            const string person1Firstname = "Anne";
            const string person1Lastname = "Jenzen";

            var person = new PersonModel {Id = 1, Firstname = person1Firstname, Lastname = person1Lastname};

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(person);

                // NOTE - can't really mock the excel export validation.. just use EPPlus if we can read the data 
                var package = new ExcelPackage(stream);

                Assert.Equal(1, package.Workbook.Worksheets.Count);

                // NOTE - EPPlus starts at 1
                var worksheet = package.Workbook.Worksheets[1];
                Assert.Equal(PersonModel.SHEET_NAME, worksheet.Name);

                // NOTE - single object get exported as KEY : VALU instea of rows
                Assert.Equal(nameof(PersonModel.Id), worksheet.Cells[1, 1].Text);
                Assert.Equal(nameof(PersonModel.Firstname), worksheet.Cells[2, 1].Text);
                Assert.Equal(nameof(PersonModel.Lastname), worksheet.Cells[3, 1].Text);

                Assert.Equal(person1Firstname, worksheet.Cells[2, 2].Text);
                Assert.Equal(person1Lastname, worksheet.Cells[3, 2].Text);
            }
        }

        [Fact]
        public void TestObjectPerson_GetParsingProperties()
        {
            var targetType = typeof(PersonModel);
            var parsingProperties = ExcelReflection.GetParsingProperties(targetType);

            Assert.Equal(3, parsingProperties.Count);

            // NOTE - GetParsingProperties must be in order of properties declared (eg for default column numbering)
            Assert.Equal(nameof(PersonModel.Id), parsingProperties[0].ColumnName);
            Assert.Equal(nameof(PersonModel.Firstname), parsingProperties[1].ColumnName);
            Assert.Equal(nameof(PersonModel.Lastname), parsingProperties[2].ColumnName);
        }

        [Fact]
        public void TestObjectPerson_GetSheetName()
        {
            var targetType = typeof(PersonModel);
            var sheetInfo = ExcelReflection.GetSheetInfo(targetType);

            Assert.Equal(PersonModel.SHEET_NAME, sheetInfo.SheetName);
        }

        [Fact]
        public void TestObjectPerson_Import()
        {
            const string person1Firstname = "Anne";
            const string person1Lastname = "Jenzen";

            var person = new PersonModel {Id = 1, Firstname = person1Firstname, Lastname = person1Lastname};

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(person);

                var importer = new ExcelModelImporter(stream);
                var outputPerson = importer.ImportModel<PersonModel>();

                Assert.Equal(person, outputPerson);
            }
        }

        [Fact]
        public void TestObjectPersonList_Export()
        {
            const string person1Firstname = "Anne";
            const string person1Lastname = "Jenzen";
            const string person2Firstname = "Julia";
            const string person2Lastname = "Jenzen";

            var persons = new[]
            {
                new PersonModel {Id = 1, Firstname = person1Firstname, Lastname = person1Lastname},
                new PersonModel {Id = 2, Firstname = person2Firstname, Lastname = person2Lastname}
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(persons);

                // NOTE - can't really mock the excel export validation.. just use EPPlus if we can read the data 
                var package = new ExcelPackage(stream);

                Assert.Equal(1, package.Workbook.Worksheets.Count);

                // NOTE - EPPlus starts at 1
                var worksheet = package.Workbook.Worksheets[1];
                Assert.Equal(PersonModel.SHEET_NAME, worksheet.Name);

                Assert.Equal(nameof(PersonModel.Id), worksheet.Cells[1, 1].Text);
                Assert.Equal(nameof(PersonModel.Firstname), worksheet.Cells[1, 2].Text);
                Assert.Equal(nameof(PersonModel.Lastname), worksheet.Cells[1, 3].Text);

                Assert.Equal(person1Firstname, worksheet.Cells[2, 2].Text);
                Assert.Equal(person1Lastname, worksheet.Cells[2, 3].Text);
                Assert.Equal(person2Firstname, worksheet.Cells[3, 2].Text);
                Assert.Equal(person2Lastname, worksheet.Cells[3, 3].Text);
            }
        }

        [Fact]
        public void TestObjectPersonList_GetSheetName()
        {
            var targetType = typeof(IList<PersonModel>);
            var sheetInfo = ExcelReflection.GetSheetInfo(targetType);

            Assert.Equal(PersonModel.SHEET_NAME, sheetInfo.SheetName);
        }

        [Fact]
        public void TestObjectPersonList_Import()
        {
            const string person1Firstname = "Anne";
            const string person1Lastname = "Jenzen";
            const string person2Firstname = "Julia";
            const string person2Lastname = "Jenzen";

            var persons = new[]
            {
                new PersonModel {Id = 1, Firstname = person1Firstname, Lastname = person1Lastname},
                new PersonModel {Id = 2, Firstname = person2Firstname, Lastname = person2Lastname}
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(persons);

                var impoter = new ExcelModelImporter(stream);
                var outputPersons = impoter.ImportModel<List<PersonModel>>();

                Assert.Equal(persons, outputPersons);
            }
        }
    }
}