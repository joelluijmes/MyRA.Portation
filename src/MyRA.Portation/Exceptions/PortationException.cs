using System;
using System.Runtime.Serialization;

namespace MyRA.Portation.Exceptions
{
    /// <summary>
    ///     Base exception for importing, exporting or parsing.
    /// </summary>
    /// <seealso cref="System.Exception" />
    public abstract class PortationException : Exception
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="PortationException" /> class.
        /// </summary>
        protected PortationException() { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="PortationException" /> class.
        /// </summary>
        /// <param name="info">
        ///     The <see cref="T:System.Runtime.Serialization.SerializationInfo"></see> that holds the serialized
        ///     object data about the exception being thrown.
        /// </param>
        /// <param name="context">
        ///     The <see cref="T:System.Runtime.Serialization.StreamingContext"></see> that contains contextual
        ///     information about the source or destination.
        /// </param>
        protected PortationException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="PortationException" /> class.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        protected PortationException(string message) : base(message) { }

        /// <summary>
        ///     Initializes a new instance of the <see cref="PortationException" /> class.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">
        ///     The exception that is the cause of the current exception, or a null reference (Nothing in
        ///     Visual Basic) if no inner exception is specified.
        /// </param>
        protected PortationException(string message, Exception innerException) : base(message, innerException) { }
    }
}
