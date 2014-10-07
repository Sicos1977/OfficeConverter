using System;

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when there is an Excel configuration problem
    /// </summary>
    public class OCExcelConfiguration : Exception
    {
        internal OCExcelConfiguration() { }

        internal OCExcelConfiguration(string message) : base(message) { }

        internal OCExcelConfiguration(string message, Exception inner) : base(message, inner) { }
    }
}