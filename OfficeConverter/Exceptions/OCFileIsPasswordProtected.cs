using System;

namespace OfficeConverter.Exceptions
{
    /// <summary>
    ///     Raised when the Microsoft Office file is password protected
    /// </summary>
    public class OCFileIsPasswordProtected : Exception
    {
        public OCFileIsPasswordProtected()
        {
        }

        public OCFileIsPasswordProtected(string message) : base(message)
        {
        }

        public OCFileIsPasswordProtected(string message, Exception inner) : base(message, inner)
        {
        }
    }
}