using System;
using System.IO;
using System.Text;

namespace OfficeConverter.Helpers
{
    internal static class Logger
    {
        #region Fields
        /// <summary>
        ///     When set then logging is written to this stream
        /// </summary>
        [ThreadStatic]
        internal static Stream LogStream;

        /// <summary>
        ///     An unique id that can be used to identify the logging of the converter when
        ///     calling the code from multiple threads and writing all the logging to the same file
        /// </summary>
        [ThreadStatic] 
        internal static string InstanceId;
        #endregion

        #region WriteToLog
        /// <summary>
        ///     Writes a line and linefeed to the <see cref="LogStream" />
        /// </summary>
        /// <param name="message">The message to write</param>
        internal static void WriteToLog(string message)
        {
            try
            {
                if (LogStream == null || !LogStream.CanWrite) return;
                var line = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss.fff") +
                           (InstanceId != null ? " - " + InstanceId : string.Empty) + " - " +
                           message + Environment.NewLine;
                var bytes = Encoding.UTF8.GetBytes(line);
                LogStream.Write(bytes, 0, bytes.Length);
                LogStream.Flush();
            }
            catch (ObjectDisposedException)
            {
                // Ignore
            }
        }
        #endregion
    }
}
