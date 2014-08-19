using System;

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file is corrupt
    /// </summary>
    public class OCFileIsCorrupt : Exception
    {
        internal OCFileIsCorrupt()
        {
        }

        internal OCFileIsCorrupt(string message) : base(message)
        {
        }

        internal OCFileIsCorrupt(string message, Exception inner) : base(message, inner)
        {
        }
    }
}