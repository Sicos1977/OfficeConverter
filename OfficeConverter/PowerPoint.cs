using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using OpenMcdf;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

/*
   Copyright 2014-2015 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace OfficeConverter
{
    /// <summary>
    /// This class is used as a placeholder for al PowerPoint related methods
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

        #region IsPasswordProtected
        /// <summary>
        /// Returns true when the binary PowerPoint file is password protected
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        internal static bool IsPasswordProtected(string fileName)
        {
            try
            {
                using (var compoundFile = new CompoundFile(fileName))
                {
                    if (compoundFile.RootStorage.TryGetStream("EncryptedPackage") != null) return true;
                    var stream = compoundFile.RootStorage.TryGetStream("Current User");
                    if (stream == null) return false;

                    using (var memoryStream = new MemoryStream(stream.GetData()))
                    using (var binaryReader = new BinaryReader(memoryStream))
                    {
                        var verAndInstance = binaryReader.ReadUInt16();
                        // ReSharper disable UnusedVariable
                        // We need to read these fields to get to the correct location in the Current User stream
                        var version = verAndInstance & 0x000FU; // first 4 bit of field verAndInstance
                        var instance = (verAndInstance & 0xFFF0U) >> 4; // last 12 bit of field verAndInstance
                        var typeCode = binaryReader.ReadUInt16();
                        var size = binaryReader.ReadUInt32();
                        var size1 = binaryReader.ReadUInt32();
                        // ReSharper restore UnusedVariable
                        var headerToken = binaryReader.ReadUInt32();

                        switch (headerToken)
                        {
                            // Not encrypted
                            case 0xE391C05F:
                                return false;

                            // Encrypted
                            case 0xF3D1C4DF:
                                return true;

                            default:
                                return false;
                        }
                    }
                }
            }
            catch (CFCorruptedFileException)
            {
                throw new OCFileIsCorrupt("The file '" + Path.GetFileName(fileName) + "' is corrupt");
            }
            catch (CFFileFormatException)
            {
                // It seems the file is just a normal Microsoft Office 2007 and up Open XML file
                return false;
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
