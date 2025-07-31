using System;

namespace ExcelAIHelper.Exceptions
{
    /// <summary>
    /// Exception thrown when an AI operation fails
    /// </summary>
    public class AiOperationException : Exception
    {
        /// <summary>
        /// Creates a new instance of AiOperationException
        /// </summary>
        public AiOperationException() : base() { }

        /// <summary>
        /// Creates a new instance of AiOperationException with a message
        /// </summary>
        /// <param name="message">The error message</param>
        public AiOperationException(string message) : base(message) { }

        /// <summary>
        /// Creates a new instance of AiOperationException with a message and inner exception
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="innerException">The inner exception</param>
        public AiOperationException(string message, Exception innerException) : base(message, innerException) { }
    }
}