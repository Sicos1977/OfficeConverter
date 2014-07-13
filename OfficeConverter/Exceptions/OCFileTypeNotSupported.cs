using System;

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file type is not supported
    /// </summary>
    public class OCFileTypeNotSupported : Exception
    {
        public OCFileTypeNotSupported()
        {
        }

        public OCFileTypeNotSupported(string message) : base(message)
        {
        }

        public OCFileTypeNotSupported(string message, Exception inner) : base(message, inner)
        {
        }
    }
}