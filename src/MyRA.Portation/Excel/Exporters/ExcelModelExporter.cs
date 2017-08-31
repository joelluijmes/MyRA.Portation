using System;
using System.Collections.Generic;
using System.IO;
using MyRA.Portation.Excel.Models;

namespace MyRA.Portation.Excel.Exporters
{
    /// <summary>
    ///     Class for exporting complex models to one or more excel sheets.
    /// </summary>
    public sealed class ExcelModelExporter
    {
        private readonly Stream stream;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelModelExporter" /> class.
        /// </summary>
        /// <param name="stream">The output stream.</param>
        /// <exception cref="ArgumentNullException">stream</exception>
        /// <exception cref="ArgumentException">Output stream must be writeable - stream</exception>
        public ExcelModelExporter(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite)
                throw new ArgumentException("Output stream must be writeable", nameof(stream));

            this.stream = stream;
        }

        /// <summary>
        ///     Exports the model.
        /// </summary>
        /// <typeparam name="TModel">The type of the model.</typeparam>
        /// <param name="model">The model.</param>
        public void ExportModel<TModel>(TModel model)
        {
            var sheetInfos = new List<ExcelSheetInfo>();

            // simple type (eg IList<TModel>) where TModel is the only sheet / rows of data
            var classSheetInfo = ExcelReflection.GetSheetInfo(typeof(TModel));
            sheetInfos.Add(classSheetInfo);

            var propertySheets = ExcelReflection.GetSheetProperties(typeof(TModel));
            sheetInfos.AddRange(propertySheets);

            using (var exporter = new ExcelSheetExporter(stream))
            {
                foreach (var property in sheetInfos)
                {
                    var value = property.GetValue(model);
                    exporter.ExportSheet(property.SheetName, value);
                }
            }
        }
    }
}
