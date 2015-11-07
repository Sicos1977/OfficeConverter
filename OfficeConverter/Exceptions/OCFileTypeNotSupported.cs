using System;
// ReSharper disable InconsistentNaming

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file type is not supported
    /// </summary>
    public class OCFileTypeNotSupported : Exception
    {
        internal OCFileTypeNotSupported() {}

        internal OCFileTypeNotSupported(string message) : base(message) {}

        internal OCFileTypeNotSupported(string message, Exception inner) : base(message, inner) {}
    }
}