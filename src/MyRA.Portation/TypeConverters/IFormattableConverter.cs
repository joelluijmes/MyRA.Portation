namespace MyRA.Portation.TypeConverters
{
    /// <summary>
    ///     Interface for type converters which accepts a format, the format gets injected at runtime during parsing.
    /// </summary>
    public interface IFormattableConverter
    {
        /// <summary>
        ///     Format used during types converting.
        /// </summary>
        /// <value>
        ///     The convert format.
        /// </value>
        string ConvertFormat { get; set; }
    }
}
