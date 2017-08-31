using System;
using System.Runtime.Serialization;

namespace MyRA.Portation.Exceptions
{
    /// <summary>
    ///     Exception thrown during exporting
    /// </summary>
    /// <seealso cref="MyRA.Portation.Exceptions.PortationException" />
    public sealed class ExportException : PortationException
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="ExportException" /> class.
        /// </summary>
        public ExportException() { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ExportException" /> class.
        /// </summary>
        /// <param name="innerException">The inner exception.</param>
        public ExportException(Exception innerException) : base("Exception while exporting", innerException) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ExportException" /> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public ExportException(string message) : base(message) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ExportException" /> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The inner exception.</param>
        public ExportException(string message, Exception innerException) : base(message, innerException) { }

        private ExportException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
