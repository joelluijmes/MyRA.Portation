using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MyRA.Portation.Exceptions;
using OfficeOpenXml;

namespace MyRA.Portation.Excel.Exporters
{
    /// <summary>
    ///     Class for exporting model to single excel sheet.
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    public sealed class ExcelSheetExporter : IDisposable
    {
        private readonly ExcelPackage package;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelSheetExporter" /> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="ArgumentNullException">stream</exception>
        /// <exception cref="ArgumentException">Output stream must be writeable - stream</exception>
        public ExcelSheetExporter(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite)
                throw new ArgumentException("Output stream must be writeable", nameof(stream));

            package = new ExcelPackage(stream);
        }

        private ExcelWorksheets Worksheets => package.Workbook.Worksheets;

        public void Dispose()
        {
            package?.Save();
            package?.Dispose();
        }

        /// <summary>
        ///     Exports data to excel sheet.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="data">The data.</param>
        public void ExportSheet(string sheetName, object data)
        {
            switch (data)
            {
                case object[] myArray:
                    var arrayType = myArray.GetType().GetElementType();
                    ExportSheetImpl(sheetName, arrayType, myArray);
                    break;

                case IEnumerable<object> list:
                    var enumerableType = list.GetType().GetGenericArguments()[0];
                    ExportSheetImpl(sheetName, enumerableType, list);
                    break;

                default:
                    ExportSheetImpl(sheetName, data.GetType(), data);
                    break;
            }
        }

        /// <summary>
        ///     Exports data to excel sheet.
        /// </summary>
        /// <typeparam name="T">Type of data</typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="data">The data.</param>
        public void ExportSheet<T>(string sheetName, T data)
        {
            ExportSheetImpl(sheetName, typeof(T), data);
        }

        /// <summary>
        ///     Exports data to excel sheet.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="data">The data.</param>
        public void ExportSheet<T>(string sheetName, IEnumerable<T> data)
        {
            ExportSheetImpl(sheetName, typeof(T), data.Cast<object>());
        }

        private void ExportSheetImpl(string sheetName, Type type, object data)
        {
            if (Worksheets[sheetName] != null)
                throw new ExportException($"Exported file already contains sheet with name {sheetName}");

            // don't create sheet for empty objects
            var parsingProperties = ExcelReflection.GetParsingProperties(type);
            if (!parsingProperties.Any())
                return;

            ExcelReflection.AutoAssignColumns(parsingProperties);
            var worksheet = Worksheets.Add(sheetName);

            // single object, so export as KEY : VALUE instead of rows of data
            foreach (var parsingProperty in parsingProperties)
            {
                if (!parsingProperty.Column.HasValue)
                    throw new InvalidOperationException($"unreachable code with property: {parsingProperty}");

                worksheet.Cells[parsingProperty.Column.Value, 1].Value = parsingProperty.ColumnName;
                worksheet.Cells[parsingProperty.Column.Value, 2].Value =
                    ExcelReflection.GetValue(parsingProperty, data);
            }
        }

        private void ExportSheetImpl(string sheetName, Type type, IEnumerable<object> data)
        {
            if (Worksheets[sheetName] != null)
                throw new ExportException($"Exported file already contains sheet with name {sheetName}");

            // don't create sheet for empty objects
            var parsingProperties = ExcelReflection.GetParsingProperties(type);
            if (!parsingProperties.Any())
                return;

            ExcelReflection.AutoAssignColumns(parsingProperties);
            var worksheet = Worksheets.Add(sheetName);

            // write column header
            foreach (var parsingProperty in parsingProperties)
            {
                if (!parsingProperty.Column.HasValue)
                    throw new InvalidOperationException($"unreachable code with property: {parsingProperty}");

                worksheet.Cells[1, parsingProperty.Column.Value].Value = parsingProperty.ColumnName;
            }

            // write the data
            // +2 : EPPlus starts at 1, and we have a column header
            var index = 2;
            foreach (var row in data)
            {
                foreach (var parsingProperty in parsingProperties)
                {
                    if (!parsingProperty.Column.HasValue)
                        throw new InvalidOperationException($"unreachable code with property: {parsingProperty}");

                    worksheet.Cells[index, parsingProperty.Column.Value].Value =
                        ExcelReflection.GetValue(parsingProperty, row);
                }

                ++index;
            }
        }
    }
}