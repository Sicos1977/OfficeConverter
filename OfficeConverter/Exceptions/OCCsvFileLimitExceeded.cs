using System;

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when a CSV file has to many rows
    /// </summary>
    public class OCCsvFileLimitExceeded : Exception
    {
        internal OCCsvFileLimitExceeded() {}

        internal OCCsvFileLimitExceeded(string message) : base(message) {}

        internal OCCsvFileLimitExceeded(string message, Exception inner) : base(message, inner) {}
    }
}