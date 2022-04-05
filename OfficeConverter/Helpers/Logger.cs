using System;
using Microsoft.Extensions.Logging;

//
// Logger.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2022 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace OfficeConverter.Helpers
{
    internal class Logger
    {
        #region Fields
        /// <summary>
        ///     When set then logging is written to this ILogger instance
        /// </summary>
        private readonly ILogger _logger;

        /// <summary>
        ///     An unique id that can be used to identify the logging of the converter when
        ///     calling the code from multiple threads and writing all the logging to the same file
        /// </summary>
        private string _instanceId;
        #endregion

        #region Constructor
        /// <summary>
        ///     Makes this object and sets all it's needed properties
        /// </summary>
        /// <param name="logger"><see cref="ILogger"/></param>
        /// <param name="instanceId"></param>
        internal Logger(ILogger logger, string instanceId)
        {
            _logger = logger;
            _instanceId = instanceId;
        }
        #endregion

        #region WriteToLog
        /// <summary>
        ///     Writes a line to the <see cref="_log" />
        /// </summary>
        /// <param name="message">The message to write</param>
        internal void WriteToLog(string message)
        {
            try
            {
                if (_logger == null) return;
                using (_logger.BeginScope(_instanceId))
                    _logger.LogInformation(message);
            }
            catch (ObjectDisposedException)
            {
                // Ignore
            }
        }
        #endregion
    }
}
