using System;

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file is password protected
    /// </summary>
    public class OCFileIsPasswordProtected : Exception
    {
        internal OCFileIsPasswordProtected() {}

        internal OCFileIsPasswordProtected(string message) : base(message) {}

        internal OCFileIsPasswordProtected(string message, Exception inner) : base(message, inner) {}
    }
}