#if VISIO_INTEROP
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using VisioInterop = Microsoft.Office.Interop.Visio;

//
// Visio.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
// Co-Author: Heinrich Elsigan
//
// Copyright (c) 2014-2025 Magic-Sessions. (www.magic-sessions.com)
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

namespace OfficeConverter
{
    /// <summary>
    ///     This class is used as a placeholder for all Visio related methods
    /// </summary>
    internal class Visio : IDisposable
    {
        #region Fields
        /// <summary>
        ///     Visio version number
        /// </summary>
        private readonly int _versionNumber;
        
        /// <summary>
        ///     <see cref="VisioInterop.ApplicationClass" />
        /// </summary>
        private VisioInterop.ApplicationClass _visio;

        /// <summary>
        ///     A <see cref="Process" /> object to Visio
        /// </summary>
        private Process _visioProcess;
        
        /// <summary>
        ///     <see cref="Logger"/>
        /// </summary>
        private readonly Logger _logger;

        /// <summary>
        ///     Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;
        #endregion

        #region Properties
        /// <summary>
        ///     Returns <c>true</c> when Visio is running
        /// </summary>
        /// <returns></returns>
        private bool IsVisioRunning
        {
            get
            {
                if (_visioProcess == null)
                    return false;

                _visioProcess.Refresh();
                return !_visioProcess.HasExited;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     This constructor checks to see if all requirements for a successful conversion are here.
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the registry could not be read to determine Visio version</exception>
        internal Visio(Logger logger)
        {
            _logger = logger;
            
            _logger?.WriteToLog("Checking what version of Visio is installed");

            try
            {
                var baseKey = Registry.ClassesRoot;
                var subKey = baseKey.OpenSubKey(@"Visio.Application\CurVer");
                if (subKey != null)
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // Visio 2003
                        case "VISIO.APPLICATION.11":
                            _versionNumber = 11;
                            break;

                        // Visio 2007
                        case "VISIO.APPLICATION.12":
                            _versionNumber = 12;
                            break;

                        // Visio 2010
                        case "VISIO.APPLICATION.14":
                            _versionNumber = 14;
                            break;

                        // Visio 2013
                        case "VISIO.APPLICATION.15":
                            _versionNumber = 15;
                            break;

                        // Visio 2016
                        case "VISIO.APPLICATION.16":
                            _versionNumber = 16;
                            break;

                        default:
                            throw new OCConfiguration("Could not determine Visio version");
                    }
                else
                    throw new OCConfiguration("Could not find registry key Visio.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check Visio version", exception);
            }
        }
        #endregion

        #region StartVisio
        /// <summary>
        ///     Starts Visio
        /// </summary>
        private void StartVisio()
        {
            if (IsVisioRunning)
            {
                _logger?.WriteToLog($"Visio is already running on PID {_visioProcess.Id}... skipped");
                return;
            }

            _logger?.WriteToLog("Starting Visio");

            _visio = new VisioInterop.ApplicationClass
            {
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
            };

            ProcessHelpers.GetWindowThreadProcessId(_visio.WindowHandle32, out var processId);
            _visioProcess = Process.GetProcessById(processId);
        
            _logger?.WriteToLog($"Visio started with process id {_visioProcess.Id}");
        }
        #endregion

        #region StopVisio
        /// <summary>
        ///     Stops Visio
        /// </summary>
        private void StopVisio()
        {
            if (IsVisioRunning)
            {
                _logger?.WriteToLog("Closing Visio gracefully");
                
                try
                {
                    _visio.Quit();
                }
                catch(Exception exception)
                {
                    _logger?.WriteToLog($"An error occurred while trying to close Visio gracefully, error '{ExceptionHelpers.GetInnerException(exception)}'");
                }

                var counter = 0;

                // Give Visio 2 seconds to close
                while (counter < 200)
                {
                    if (!IsVisioRunning)
                    {
                        _logger?.WriteToLog("Visio closed gracefully");
                        break;
                    }

                    counter++;
                    Thread.Sleep(10);
                }

                if (IsVisioRunning)
                {
                    _logger?.WriteToLog($"Visio did not close gracefully, closing it by killing it's process on id {_visioProcess.Id}");
                    _visioProcess.Kill();
                    _visioProcess = null;
                    _logger?.WriteToLog("Visio process killed");
                }
            }
            else
                _logger?.WriteToLog($"Visio {(_visioProcess != null ? $"with process id {_visioProcess.Id} " : string.Empty)}already exited");

            if (_visio != null)
            {
                Marshal.ReleaseComObject(_visio);
                _visio = null;
            }

            _visioProcess = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts a Visio document to PDF
        /// </summary>
        /// <param name="inputFile">The Visio input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        internal void Convert(string inputFile, string outputFile)
        {
            DeleteResiliencyKeys();

            VisioInterop.Document document = null;

            try
            {
                StartVisio();

                document = OpenDocument(inputFile, false);

                _logger?.WriteToLog($"Exporting document to PDF file '{outputFile}'");
                document.ExportAsFixedFormat(VisioInterop.VisFixedFormatTypes.visFixedFormatPDF, 
                    outputFile, 
                    VisioInterop.VisDocExIntent.visDocExIntentScreen, 
                    VisioInterop.VisPrintOutRange.visPrintAll);
                _logger?.WriteToLog("Document exported to PDF");
            }
            catch (Exception exception)
            {
                _logger?.WriteToLog($"An error occurred, error '{ExceptionHelpers.GetInnerException(exception)}'");
                StopVisio();
                throw;
            }
            finally
            {
                CloseDocument(document);
            }
        }
        #endregion

        #region OpenDocument
        /// <summary>
        ///     Opens the <paramref name="inputFile" /> and returns it as an <see cref="VisioInterop.Document" /> object
        /// </summary>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile" /> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">
        ///     Raised when the <paramref name="inputFile" /> is corrupt and can't be opened in
        ///     repair mode
        /// </exception>
        private VisioInterop.Document OpenDocument(string inputFile, bool repairMode)
        {
            try
            {
                return _visio.Documents.Open(inputFile);
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt($"The file '{Path.GetFileName(inputFile)}' seems to be corrupt, error: {ExceptionHelpers.GetInnerException(exception)}");

                return OpenDocument(inputFile, true);
            }
        }
        #endregion

        #region CloseDocument
        /// <summary>
        ///     Closes the opened document and releases any allocated resources
        /// </summary>
        private void CloseDocument(VisioInterop.Document document)
        {
            if (document == null) return;
            _logger?.WriteToLog("Closing document");
            document.Close();
            Marshal.ReleaseComObject(document);
            _logger?.WriteToLog("Document closed");
        }
        #endregion

        #region DeleteResiliencyKeys
        /// <summary>
        ///     This method will delete the automatic created Resiliency key. Visio uses this registry key
        ///     to make entries to corrupted documents. If there are to many entries under this key Visio will
        ///     get slower and slower to start. To prevent this we just delete this key when it exists
        /// </summary>
        private void DeleteResiliencyKeys()
        {
            _logger?.WriteToLog("Deleting Visio resiliency keys from the registry");

            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Resiliency\DocumentRecovery
                var key = $@"Software\Microsoft\Office\{_versionNumber}.0\Visio\Resiliency";

                if (Registry.CurrentUser.OpenSubKey(key, false) != null)
                {
                    Registry.CurrentUser.DeleteSubKeyTree(key);
                    _logger?.WriteToLog("Resiliency keys deleted");
                }
                else
                    _logger?.WriteToLog("There are no keys to delete");
            }
            catch (Exception exception)
            {
                _logger?.WriteToLog($"Failed to delete resiliency keys, error: {ExceptionHelpers.GetInnerException(exception)}");
            }
        }
        #endregion

        #region Dispose
        /// <summary>
        ///     Disposes the running <see cref="_visio" />
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            StopVisio();
            _disposed = true;
        }
        #endregion
    }
}
#endif
