using System;
using System.Runtime.Serialization;

namespace MyRA.Portation.Exceptions
{
    /// <summary>
    ///     Exception thrown when parsing
    /// </summary>
    /// <seealso cref="PortationException" />
    public sealed class ParserException : PortationException
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="ParserException" /> class.
        /// </summary>
        public ParserException() : base("Exception during parsing.")
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ParserException" /> class.
        /// </summary>
        /// <param name="innerException">The inner exception.</param>
        public ParserException(Exception innerException) : base("Exception during parsing.", innerException)
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ParserException" /> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public ParserException(string message) : base(message)
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="ParserException" /> class.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">
        ///     The exception that is the cause of the current exception, or a null reference (Nothing in
        ///     Visual Basic) if no inner exception is specified.
        /// </param>
        public ParserException(string message, Exception innerException) : base(message, innerException)
        {
        }

        private ParserException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}