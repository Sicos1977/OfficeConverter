using System;

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file is corrupt
    /// </summary>
    public class OCFileIsCorrupt : Exception
    {
        public OCFileIsCorrupt()
        {
        }

        public OCFileIsCorrupt(string message) : base(message)
        {
        }

        public OCFileIsCorrupt(string message, Exception inner) : base(message, inner)
        {
        }
    }
}