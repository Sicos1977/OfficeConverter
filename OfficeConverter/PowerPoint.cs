using System;
using System.IO;
using System.Runtime.InteropServices;
using CompoundFileStorage;
using CompoundFileStorage.Exceptions;
using Microsoft.Office.Core;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace OfficeConverter
{
    /// <summary>
    /// This class is used as a placeholder for al PowerPoint related methods
    /// </summary>
    internal static class PowerPoint
    {
        #region Convert
        /// <summary>
        /// Converts a PowerPoint document to PDF
        /// </summary>
        /// <param name="inputFile">The PowerPoint input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        internal static void Convert(string inputFile, string outputFile)
        {
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

        #region FileIsPasswordProtected
        /// <summary>
        /// Returns true when the binary PowerPoint file is password protected
        /// </summary>
        /// <param name="inputFile">The PowerPoint file to check</param>
        /// <returns></returns>
        internal static bool FileIsPasswordProtected(string inputFile)
        {
            try
            {
                using (var compoundFile = new CompoundFile(inputFile))
                {
                    if (compoundFile.RootStorage.ExistsStream("EncryptedPackage")) return true;
                    if (!compoundFile.RootStorage.ExistsStream("Current User")) return false;
                    var stream = compoundFile.RootStorage.GetStream("Current User") as CFStream;
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
                throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' is corrupt");
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
        private static PowerPointInterop.Presentation Open(PowerPointInterop._Application powerPoint,
                                                           string inputFile,
                                                           bool repairMode)
        {
            try
            {
                return powerPoint.Presentations.Open(inputFile, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' seems to be corrupt, error: " + ExceptionHelpers.GetInnerException(exception));

                return Open(powerPoint, inputFile, true);
            }
        }
        #endregion
    }
}
