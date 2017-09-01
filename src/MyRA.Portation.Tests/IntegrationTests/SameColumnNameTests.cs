using System.Collections.Generic;
using System.IO;
using MyRA.Portation.Excel.Exporters;
using MyRA.Portation.Excel.Importers;
using MyRA.Portation.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace MyRA.Portation.Tests.IntegrationTests
{
    public sealed class SameColumnNameTests
    {
        [Fact]
        public void TestObjectSameName_Export()
        {
            const string firstname1 = "Julia";
            const string lastname1 = "Jansen";
            const string firstname2 = "Anne";
            const string lastname2 = "Kamphuis";

            var models = new[]
            {
                new SameColumnNameModel {Firstname = firstname1, Lastname = lastname1},
                new SameColumnNameModel {Firstname = firstname2, Lastname = lastname2}
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(models);

                // NOTE - can't really mock the excel export validation.. just use EPPlus if we can read the data 
                var package = new ExcelPackage(stream);

                // two sheets, one for items one for person
                Assert.Equal(1, package.Workbook.Worksheets.Count);

                // NOTE - EPPlus starts at 1
                
                var worksheet = package.Workbook.Worksheets[1];
                Assert.Equal(SameColumnNameModel.COLUMN_NAME, worksheet.Cells[1, 1].Text);
                Assert.Equal(SameColumnNameModel.COLUMN_NAME, worksheet.Cells[1, 2].Text);
                Assert.Equal(firstname1, worksheet.Cells[2, 1].Text);
                Assert.Equal(lastname1, worksheet.Cells[2, 2].Text);
                Assert.Equal(firstname2, worksheet.Cells[3, 1].Text);
                Assert.Equal(lastname2, worksheet.Cells[3, 2].Text);
            }
        }

        [Fact]
        public void TestObjectSameName_Import()
        {
            const string firstname1 = "Julia";
            const string lastname1 = "Jansen";
            const string firstname2 = "Anne";
            const string lastname2 = "Kamphuis";

            var models = new[]
            {
                new SameColumnNameModel {Firstname = firstname1, Lastname = lastname1},
                new SameColumnNameModel {Firstname = firstname2, Lastname = lastname2}
            };

            using (var stream = new MemoryStream())
            {
                var exporter = new ExcelModelExporter(stream);
                exporter.ExportModel(models);

                var importer = new ExcelModelImporter(stream);
                var outputModels = importer.ImportModel<List<SameColumnNameModel>>();

                Assert.Equal<SameColumnNameModel>(models, outputModels);
            }
        }
    }
}
