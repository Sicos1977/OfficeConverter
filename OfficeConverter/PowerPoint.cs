using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

//
// PowerPoint.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2021 Magic-Sessions. (www.magic-sessions.com)
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
    ///     This class is used as a placeholder for all PowerPoint related methods
    /// </summary>
    internal class PowerPoint : IDisposable
    {
        #region Fields
        /// <summary>
        ///     PowerPoint version number
        /// </summary>
        private readonly int _versionNumber;
        
        /// <summary>
        ///     <see cref="PowerPointInterop.ApplicationClass" />
        /// </summary>
        private PowerPointInterop.ApplicationClass _powerPoint;

        /// <summary>
        ///     A <see cref="Process" /> object to PowerPoint
        /// </summary>
        private Process _powerPointProcess;

        /// <summary>
        ///     Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;
        #endregion

        #region Properties
        /// <summary>
        ///     Returns <c>true</c> when PowerPoint is running
        /// </summary>
        /// <returns></returns>
        private bool IsPowerPointRunning
        {
            get
            {
                if (_powerPointProcess == null)
                    return false;

                _powerPointProcess.Refresh();
                return !_powerPointProcess.HasExited;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     This constructor checks to see if all requirements for a successful conversion are here.
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the registry could not be read to determine PowerPoint version</exception>
        internal PowerPoint()
        {
            Logger.WriteToLog("Checking what version of PowerPoint is installed");

            try
            {
                var baseKey = Registry.ClassesRoot;
                var subKey = baseKey.OpenSubKey(@"PowerPoint.Application\CurVer");
                if (subKey != null)
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // PowerPoint 2003
                        case "POWERPOINT.APPLICATION.11":
                            _versionNumber = 11;
                            break;

                        // PowerPoint 2007
                        case "POWERPOINT.APPLICATION.12":
                            _versionNumber = 12;
                            break;

                        // PowerPoint 2010
                        case "POWERPOINT.APPLICATION.14":
                            _versionNumber = 14;
                            break;

                        // PowerPoint 2013
                        case "POWERPOINT.APPLICATION.15":
                            _versionNumber = 15;
                            break;

                        // PowerPoint 2016
                        case "POWERPOINT.APPLICATION.16":
                            _versionNumber = 16;
                            break;

                        default:
                            throw new OCConfiguration("Could not determine PowerPoint version");
                    }
                else
                    throw new OCConfiguration("Could not find registry key PowerPoint.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check PowerPoint version", exception);
            }
        }
        #endregion

        #region StartPowerPoint
        /// <summary>
        ///     Starts PowerPoint
        /// </summary>
        private void StartPowerPoint()
        {
            if (IsPowerPointRunning)
            {
                Logger.WriteToLog($"Powerpoint is already running on PID {_powerPointProcess.Id}... skipped");
                return;
            }

            Logger.WriteToLog("Starting PowerPoint");

            _powerPoint = new PowerPointInterop.ApplicationClass
            {
                DisplayAlerts = PowerPointInterop.PpAlertLevel.ppAlertsNone,
                DisplayDocumentInformationPanel = false,
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
            };

            ProcessHelpers.GetWindowThreadProcessId(_powerPoint.HWND, out var processId);
            _powerPointProcess = Process.GetProcessById(processId);
        
            Logger.WriteToLog($"PowerPoint started with process id {_powerPointProcess.Id}");
        }
        #endregion

        #region StopPowerPoint
        /// <summary>
        ///     Stops PowerPoint
        /// </summary>
        private void StopPowerPoint()
        {
            if (IsPowerPointRunning)
            {
                Logger.WriteToLog("Stopping PowerPoint");
                
                try
                {
                    _powerPoint.Quit();
                }
                catch(Exception exception)
                {
                    Logger.WriteToLog($"PowerPoint did not shutdown gracefully, exception: {ExceptionHelpers.GetInnerException(exception)}");
                }

                var counter = 0;

                // Give PowerPoint 2 seconds to close
                while (counter < 200)
                {
                    if (!IsPowerPointRunning) break;
                    counter++;
                    Thread.Sleep(10);
                }

                if (IsPowerPointRunning)
                {
                    Logger.WriteToLog($"PowerPoint did not shutdown gracefully in 2 seconds ... killing it on process id {_powerPointProcess.Id}");
                    _powerPointProcess.Kill();
                    _powerPointProcess = null;
                    Logger.WriteToLog("PowerPoint process killed");
                }
                else
                    Logger.WriteToLog("PowerPoint stopped");
            }
            else
                Logger.WriteToLog($"PowerPoint {(_powerPointProcess != null ? $"with process id {_powerPointProcess.Id} " : string.Empty)}already exited");

            if (_powerPoint != null)
            {
                Marshal.ReleaseComObject(_powerPoint);
                _powerPoint = null;
            }

            _powerPointProcess = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts a PowerPoint document to PDF
        /// </summary>
        /// <param name="inputFile">The PowerPoint input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        internal void Convert(string inputFile, string outputFile)
        {
            DeleteResiliencyKeys();

            PowerPointInterop.Presentation presentation = null;

            try
            {
                StartPowerPoint();

                presentation = OpenPresentation(inputFile, false);

                Logger.WriteToLog($"Exporting presentation to PDF file '{outputFile}'");
                presentation.ExportAsFixedFormat(outputFile, PowerPointInterop.PpFixedFormatType.ppFixedFormatTypePDF);
                Logger.WriteToLog("Presentation exported to PDF");
            }
            catch (Exception)
            {
                StopPowerPoint();
                throw;
            }
            finally
            {
                ClosePresentation(presentation);
            }
        }
        #endregion

        #region OpenPresentation
        /// <summary>
        ///     Opens the <paramref name="inputFile" /> and returns it as an <see cref="PowerPointInterop.Presentation" /> object
        /// </summary>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile" /> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">
        ///     Raised when the <paramref name="inputFile" /> is corrupt and can't be opened in
        ///     repair mode
        /// </exception>
        private PowerPointInterop.Presentation OpenPresentation(string inputFile, bool repairMode)
        {
            try
            {
                return _powerPoint.Presentations.Open(inputFile, MsoTriState.msoTrue, MsoTriState.msoTrue,
                    MsoTriState.msoFalse);
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt($"The file '{Path.GetFileName(inputFile)}' seems to be corrupt, error: {ExceptionHelpers.GetInnerException(exception)}");

                return OpenPresentation(inputFile, true);
            }
        }
        #endregion

        #region ClosePresentation
        /// <summary>
        ///     Closes the opened presentation and releases any allocated resources
        /// </summary>
        private void ClosePresentation(PowerPointInterop._Presentation presentation)
        {
            if (presentation == null) return;
            Logger.WriteToLog("Closing presentation");
            presentation.Saved = MsoTriState.msoFalse;
            presentation.Close();
            Marshal.ReleaseComObject(presentation);
            Logger.WriteToLog("Presentation closed");
        }
        #endregion

        #region DeleteResiliencyKeys
        /// <summary>
        ///     This method will delete the automatic created Resiliency key. PowerPoint uses this registry key
        ///     to make entries to corrupted presentations. If there are to many entries under this key PowerPoint will
        ///     get slower and slower to start. To prevent this we just delete this key when it exists
        /// </summary>
        private void DeleteResiliencyKeys()
        {
            Logger.WriteToLog("Deleting PowerPoint resiliency keys from the registry");

            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\PowerPoint\Resiliency\DocumentRecovery
                var key = $@"Software\Microsoft\Office\{_versionNumber}.0\PowerPoint\Resiliency";

                if (Registry.CurrentUser.OpenSubKey(key, false) != null)
                {
                    Registry.CurrentUser.DeleteSubKeyTree(key);
                    Logger.WriteToLog("Resiliency keys deleted");
                }
                else
                    Logger.WriteToLog("There are no keys to delete");
            }
            catch (Exception exception)
            {
                Logger.WriteToLog($"Failed to delete resiliency keys, error: {ExceptionHelpers.GetInnerException(exception)}");
            }
        }
        #endregion

        #region Dispose
        /// <summary>
        ///     Disposes the running <see cref="_powerPoint" />
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            StopPowerPoint();
        }
        #endregion
    }
}