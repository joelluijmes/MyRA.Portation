using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using MyRA.Portation.Excel.Models;
using MyRA.Portation.Exceptions;
using OfficeOpenXml;

namespace MyRA.Portation.Excel.Importers
{
    /// <summary>
    ///     Class for importing model from single excel sheet.
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    public sealed class ExcelSheetImporter : IDisposable
    {
        private readonly ExcelPackage package;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelSheetImporter" /> class.
        /// </summary>
        /// <param name="stream">The input stream.</param>
        /// <exception cref="ArgumentNullException">stream</exception>
        /// <exception cref="ArgumentException">Input stream must be readable - stream</exception>
        /// <exception cref="ImportException">
        ///     Exception during creating ExcelPackage. Hint: is the input stream a valid .xlsx file? Note that .xls is NOT
        ///     supported by EPPlus.
        ///     or
        ///     Exception during creating ExcelPackage.
        /// </exception>
        public ExcelSheetImporter(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead)
                throw new ArgumentException("Input stream must be readable", nameof(stream));

            try
            {
                package = new ExcelPackage(stream);
            }
            catch (COMException exception)
            {
                throw new ImportException(
                    "Exception during creating ExcelPackage. Hint: is the input stream a valid .xlsx file? Note that .xls is NOT supported by EPPlus.",
                    exception);
            }
            catch (Exception exception)
            {
                throw new ImportException("Exception during creating ExcelPackage.", exception);
            }
        }

        private ExcelWorksheets Worksheets => package.Workbook.Worksheets;

        public void Dispose()
        {
            package?.Dispose();
        }

        /// <summary>
        ///     Imports model from sheet.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        public T ImportSheet<T>(string sheetName)
            where T : new()
        {
            return (T) ImportSheet(sheetName, typeof(T));
        }

        /// <summary>
        ///     Imports model from sheet.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        public object ImportSheet(string sheetName, Type type)
        {
            var enumerableType = ExcelReflection.GetGenericTypeFromEnumerableOrDefault(type);
            return enumerableType == null
                ? ImportSheetImpl(sheetName, type)
                : ImportSheetAsListImpl(sheetName, enumerableType);
        }

        /// <summary>
        ///     Import model as a List from sheet.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns></returns>
        public IList<T> ImportSheetAsList<T>(string sheetName)
        {
            return ImportSheetAsListImpl(sheetName, typeof(T)).Cast<T>().ToList();
        }

        private static void FixColumns(ExcelRange headerRange, IList<ExcelPropertyInfo> parsingProperties,
            ColomnOrientation orientation)
        {
            var startRow = headerRange.Start.Row;
            var startCol = headerRange.Start.Column;

            var start = orientation == ColomnOrientation.Horizontal
                ? headerRange.Start.Column
                : headerRange.Start.Row;
            var end = orientation == ColomnOrientation.Horizontal
                ? headerRange.End.Column
                : headerRange.End.Row;

            var takenColumns = new HashSet<int>();
            foreach (var property in parsingProperties)
            {
                int? column = null;
                if (property.Attribute.Column.HasValue)
                {
                    if (property.Attribute.Column.Value > end || property.Attribute.Column.Value < start)
                        throw new ImportException($"Parsing property for {property} is out of range {start} - {end}");

                    column = property.Attribute.Column;
                }
                else if (!string.IsNullOrEmpty(property.ColumnName))
                {
                    for (var xy = start; xy <= end; ++xy)
                    {
                        var cell = orientation == ColomnOrientation.Horizontal
                            ? headerRange[startRow, xy].Text
                            : headerRange[xy, startCol].Text;

                        if (!string.Equals(cell, property.ColumnName, StringComparison.CurrentCultureIgnoreCase))
                            continue;

                        // NOTE - we keep on iterating, even if we have found the column because we are going to throw exception if we find duplicate
                        if (column.HasValue)
                            throw new ImportException(
                                $"Could not parse columns; column {property} was already found at {column}");
                        if (takenColumns.Contains(xy))
                            throw new ImportException(
                                $"Could not parse columns; column {property} is already been taken");

                        column = xy;
                    }
                }
                else
                {
                    throw new ImportException(
                        "Could not parse columns; current object does not have harcoded Column or ColumnName");
                }

                property.Column = column;
                takenColumns.Add(column.Value);
            }

            var missingProperties = parsingProperties
                .Where(p => !p.Column.HasValue)
                .ToArray();

            if (!missingProperties.Any())
                return;

            var missingColumns = missingProperties
                .Select(p => p.ColumnName)
                .Aggregate((cur, acc) => $"{cur}, {acc}");

            throw new ImportException($"Sheet does not contain columns for: {missingColumns}");
        }

        private IList ImportSheetAsListImpl(string sheetName, Type type)
        {
            var list = (IList) Activator.CreateInstance(typeof(List<>).MakeGenericType(type));

            // ignore empty objects, just give the constructed list
            var parsingProperties = ExcelReflection.GetParsingProperties(type);
            if (!parsingProperties.Any())
                return list;

            var worksheet = Worksheets[sheetName];
            if (worksheet == null)
                throw new ImportException($"Import file missing sheet with name {sheetName}");

            // row of objects, column at first 
            var headerRange = worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                worksheet.Dimension.Start.Row, worksheet.Dimension.End.Column];
            FixColumns(headerRange, parsingProperties, ColomnOrientation.Horizontal);

            // skip the header
            var startRow = worksheet.Dimension.Start.Row + 1;
            var endRow = worksheet.Dimension.End.Row;

            for (var row = startRow; row <= endRow; ++row)
            {
                var item = Activator.CreateInstance(type);
                foreach (var parsingProperty in parsingProperties)
                {
                    if (!parsingProperty.Column.HasValue)
                        throw new InvalidOperationException($"unreachable code with property: {parsingProperty}");

                    var value = worksheet.Cells[row, parsingProperty.Column.Value].Text;
                    ExcelReflection.SetConvertedValue(parsingProperty, item, value);
                }

                list.Add(item);
            }

            return list;
        }

        private object ImportSheetImpl(string sheetName, Type type)
        {
            var item = Activator.CreateInstance(type);

            // ignore empty objects, just give the constructed item for them 
            var parsingProperties = ExcelReflection.GetParsingProperties(type);
            if (!parsingProperties.Any())
                return item;

            var worksheet = Worksheets[sheetName];
            if (worksheet == null)
                throw new ImportException($"Import file missing sheet with name {sheetName}");

            // single object, import as KEY : VALUE instead of multiple rows
            var headerRange = worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                worksheet.Dimension.End.Row, worksheet.Dimension.Start.Column];
            FixColumns(headerRange, parsingProperties, ColomnOrientation.Vertical);

            foreach (var parsingProperty in parsingProperties)
            {
                if (!parsingProperty.Column.HasValue)
                    throw new InvalidOperationException($"unreachable code with property: {parsingProperty}");

                var value = worksheet.Cells[parsingProperty.Column.Value, worksheet.Dimension.Start.Column + 1].Text;
                ExcelReflection.SetConvertedValue(parsingProperty, item, value);
            }

            return item;
        }

        private enum ColomnOrientation
        {
            Vertical,
            Horizontal
        }
    }
}