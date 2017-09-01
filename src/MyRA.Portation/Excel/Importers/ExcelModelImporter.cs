using System;
using System.Collections.Generic;
using System.IO;
using MyRA.Portation.Excel.Models;

namespace MyRA.Portation.Excel.Importers
{
    /// <summary>
    ///     Class for importing complex models from multi-sheet excel.
    /// </summary>
    public sealed class ExcelModelImporter
    {
        private readonly Stream stream;

        /// <summary>
        ///     Initializes a new instance of the <see cref="ExcelModelImporter" /> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <exception cref="ArgumentNullException">stream</exception>
        /// <exception cref="ArgumentException">Input stream must be readable - stream</exception>
        public ExcelModelImporter(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead)
                throw new ArgumentException("Input stream must be readable", nameof(stream));

            this.stream = stream;
        }

        /// <summary>
        ///     Import complex model from excel sheet(s).
        /// </summary>
        /// <typeparam name="TModel">The type of the model.</typeparam>
        /// <returns></returns>
        public TModel ImportModel<TModel>()
            where TModel : new()
        {
            var sheetInfos = new List<ExcelSheetInfo>();

            // simple type (eg IList<TModel>) where TModel is the only sheet / rows of data
            var classSheetInfo = ExcelReflection.GetSheetInfo(typeof(TModel));
            sheetInfos.Add(classSheetInfo);

            var propertySheets = ExcelReflection.GetSheetProperties(typeof(TModel));
            sheetInfos.AddRange(propertySheets);

            object model = new TModel();
            using (var importer = new ExcelSheetImporter(stream))
            {
                foreach (var sheetInfo in sheetInfos)
                {
                    var value = importer.ImportSheet(sheetInfo.SheetName, sheetInfo.Type);
                    sheetInfo.SetValue(ref model, value);
                }
            }

            return (TModel) model;
        }
    }
}