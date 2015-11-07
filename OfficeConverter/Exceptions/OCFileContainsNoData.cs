using System;
// ReSharper disable InconsistentNaming

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file contains no actual data
    /// </summary>
    public class OCFileContainsNoData : Exception
    {
        internal OCFileContainsNoData() { }

        internal OCFileContainsNoData(string message) : base(message) { }

        internal OCFileContainsNoData(string message, Exception inner) : base(message, inner) { }
    }
}