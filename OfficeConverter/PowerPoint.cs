using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
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
// Copyright (c) 2014-2018 Magic-Sessions. (www.magic-sessions.com)
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
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace OfficeConverter
{
    /// <summary>
    /// This class is used as a placeholder for all PowerPoint related methods
    /// </summary>
    internal static class PowerPoint
    {
        #region Fields
        /// <summary>
        /// Excel version number
        /// </summary>
        private static readonly int VersionNumber;
        #endregion

        #region Constructor
        /// <summary>
        /// This constructor is called the first time when the <see cref="Convert"/> or
        /// <see cref="FileIsPasswordProtected"/> method is called. Some checks are done to
        /// see if all requirements for a succesfull conversion are there.
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the registry could not be read to determine PowerPoint version</exception>
        static PowerPoint()
        {
            try
            {
                var baseKey = Registry.ClassesRoot;
                var subKey = baseKey.OpenSubKey(@"PowerPoint.Application\CurVer");
                if (subKey != null)
                {
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // PowerPoint 2003
                        case "POWERPOINT.APPLICATION.11":
                            VersionNumber = 11;
                            break;

                        // PowerPoint 2007
                        case "POWERPOINT.APPLICATION.12":
                            VersionNumber = 12;
                            break;

                        // PowerPoint 2010
                        case "POWERPOINT.APPLICATION.14":
                            VersionNumber = 14;
                            break;

                        // PowerPoint 2013
                        case "POWERPOINT.APPLICATION.15":
                            VersionNumber = 15;
                            break;

                        // PowerPoint 2016
                        case "POWERPOINT.APPLICATION.16":
                            VersionNumber = 16;
                            break;

                        default:
                            throw new OCConfiguration("Could not determine PowerPoint version");
                    }
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

        #region Convert
        /// <summary>
        /// Converts a PowerPoint document to PDF
        /// </summary>
        /// <param name="inputFile">The PowerPoint input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        internal static void Convert(string inputFile, string outputFile)
        {
            DeleteAutoRecoveryFiles();

            PowerPointInterop.ApplicationClass powerPoint = null;
            PowerPointInterop.Presentation presentation = null;

            try
            {
                powerPoint = new PowerPointInterop.ApplicationClass
                {
                    DisplayAlerts = PowerPointInterop.PpAlertLevel.ppAlertsNone,
                    DisplayDocumentInformationPanel = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                };

                presentation = Open(powerPoint, inputFile, false);
                presentation.ExportAsFixedFormat(outputFile, PowerPointInterop.PpFixedFormatType.ppFixedFormatTypePDF);
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Saved = MsoTriState.msoFalse;
                    presentation.Close();
                    Marshal.ReleaseComObject(presentation);
                }

                if (powerPoint != null)
                {
                    powerPoint.Quit();
                    Marshal.ReleaseComObject(powerPoint);
                }
            }
        }
        #endregion

        #region Open
        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="PowerPointInterop.Presentation"/> object
        /// </summary>
        /// <param name="powerPoint">The <see cref="PowerPointInterop.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">Raised when the <paramref name="inputFile"/> is corrupt and can't be opened in repair mode</exception>
        private static PowerPointInterop.Presentation Open(PowerPointInterop._Application powerPoint, string inputFile, bool repairMode)
        {
            try
            {
                return powerPoint.Presentations.Open(inputFile, MsoTriState.msoTrue, MsoTriState.msoTrue,
                    MsoTriState.msoFalse);
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' seems to be corrupt, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return Open(powerPoint, inputFile, true);
            }
        }
        #endregion

        #region DeleteAutoRecoveryFiles
        /// <summary>
        /// This method will delete the automatic created Resiliency key. PowerPoint uses this registry key  
        /// to make entries to corrupted presentations. If there are to many entries under this key PowerPoint will
        /// get slower and slower to start. To prevent this we just delete this key when it exists
        /// </summary>
        private static void DeleteAutoRecoveryFiles()
        {
            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\PowerPoint\Resiliency\DocumentRecovery
                var version = string.Empty;

                switch (VersionNumber)
                {
                    // Word 2003
                    case 11:
                        version = "11.0";
                        break;

                    // Word 2017
                    case 12:
                        version = "12.0";
                        break;

                    // Word 2010
                    case 14:
                        version = "14.0";
                        break;

                    // Word 2013
                    case 15:
                        version = "15.0";
                        break;

                    // Word 2016
                    case 16:
                        version = "16.0";
                        break;
                }

                var key = @"Software\Microsoft\Office\" + version + @"\PowerPoint\Resiliency";

                if (Registry.CurrentUser.OpenSubKey(key, false) != null)
                    Registry.CurrentUser.DeleteSubKeyTree(key);
            }
            catch (Exception exception)
            {
                EventLog.WriteEntry("OfficeConverter", ExceptionHelpers.GetInnerException(exception), EventLogEntryType.Error);
            }
        }
        #endregion
    }
}
