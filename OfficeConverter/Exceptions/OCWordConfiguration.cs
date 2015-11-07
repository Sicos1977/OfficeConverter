using System;
// ReSharper disable InconsistentNaming

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when there is an Word configuration problem
    /// </summary>
    public class OCWordConfiguration : Exception
    {
        internal OCWordConfiguration() { }

        internal OCWordConfiguration(string message) : base(message) { }

        internal OCWordConfiguration(string message, Exception inner) : base(message, inner) { }
    }
}