using System;
using System.Runtime.Serialization;

namespace MyRA.Portation.Exceptions
{
    /// <summary>
    ///     Exception thrown when importing.
    /// </summary>
    /// <seealso cref="PortationException" />
    public sealed class ImportException : PortationException
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="ImportException" /> class.
        /// </summary>
        public ImportException()
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ImportException" /> class.
        /// </summary>
        /// <param name="innerException">The inner exception.</param>
        public ImportException(Exception innerException) : base("Exception while importing", innerException)
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ImportException" /> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public ImportException(string message) : base(message)
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ImportException" /> class.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">
        ///     The exception that is the cause of the current exception, or a null reference (Nothing in
        ///     Visual Basic) if no inner exception is specified.
        /// </param>
        public ImportException(string message, Exception innerException) : base(message, innerException)
        {
        }

        private ImportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}