using System;
using System.IO;
using System.Runtime.InteropServices;
using CompoundFileStorage;
using CompoundFileStorage.Exceptions;
using Microsoft.Office.Core;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using WordInterop = Microsoft.Office.Interop.Word;

namespace OfficeConverter
{
    /// <summary>
    /// This class is used as a placeholder for al Word related methods
    /// </summary>
    internal static class Word
    {
        #region Convert
        /// <summary>
        /// Converts a Word document to PDF
        /// </summary>
        /// <param name="inputFile">The Word input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        internal static void Convert(string inputFile, string outputFile)
        {
            WordInterop.ApplicationClass word = null;
            WordInterop.DocumentClass document = null;

            try
            {
                word = new WordInterop.ApplicationClass
                {
                    ScreenUpdating = false,
                    DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                };

                word.Options.UpdateLinksAtOpen = false;
                word.Options.ConfirmConversions = false;
                word.Options.SaveInterval = 0;
                word.Options.SaveNormalPrompt = false;
                word.Options.SavePropertiesPrompt = false;
                word.Options.AllowReadingMode = false;
                word.Options.WarnBeforeSavingPrintingSendingMarkup = false;
                word.Options.UpdateFieldsAtPrint = false;
                word.Options.UpdateLinksAtOpen = false;
                word.Options.UpdateLinksAtPrint = false;
                
                document = (WordInterop.DocumentClass)Open(word, inputFile, false);

                // Do not remove this line!!
                // This is yet another solution to a weird Office problem. Sometimes there
                // are Word documents with images in it that take some time to load. When
                // we remove the line below the ExportAsFixedFormat method will be called 
                // before the images are loaded thus resulting in an unendless loop somewhere
                // in this method.
                // ReSharper disable once UnusedVariable
                var count = document.ComputeStatistics(WordInterop.WdStatistic.wdStatisticPages);

                word.DisplayAutoCompleteTips = false;
                word.DisplayScreenTips = false;
                word.DisplayStatusBar = false;

                document.ExportAsFixedFormat(outputFile, WordInterop.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                if (document != null)
                {
                    document.Saved = true;
                    document.Close(false);
                    Marshal.ReleaseComObject(document);
                }

                if (word != null)
                {
                    word.Quit(false);
                    Marshal.ReleaseComObject(word);
                }
            }
        }
        #endregion

        #region FileIsPasswordProtected
        /// <summary>
        /// Returns true when the Word file is password protected
        /// </summary>
        /// <param name="inputFile">The Word file to check</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">Raised when the file is corrupt</exception>
        internal static bool FileIsPasswordProtected(string inputFile)
        {
            try
            {
                using (var compoundFile = new CompoundFile(inputFile))
                {
                    if (compoundFile.RootStorage.ExistsStream("EncryptedPackage")) return true;
                    if (!compoundFile.RootStorage.ExistsStream("WordDocument"))
                        throw new OCFileIsCorrupt("Could not find the WordDocument stream in the file '" +
                                                  compoundFile.FileName + "'");

                    var stream = compoundFile.RootStorage.GetStream("WordDocument") as CFStream;
                    if (stream == null) return false;

                    var bytes = stream.GetData();
                    using (var memoryStream = new MemoryStream(bytes))
                    using (var binaryReader = new BinaryReader(memoryStream))
                    {
                        //http://msdn.microsoft.com/en-us/library/dd944620%28v=office.12%29.aspx
                        // The bit that shows if the file is encrypted is in the 11th and 12th byte so we 
                        // need to skip the first 10 bytes
                        binaryReader.ReadBytes(10);

                        // Now we read the 2 bytes that we need
                        var pnNext = binaryReader.ReadUInt16();
                        //(value & mask) == mask)

                        // The bit that tells us if the file is encrypted
                        return (pnNext & 0x0100) == 0x0100;
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
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="WordInterop.Document"/> object
        /// </summary>
        /// <param name="word">The <see cref="WordInterop.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        private static WordInterop.Document Open(WordInterop._Application word,
                                                string inputFile,
                                                bool repairMode)
        {
            try
            {
                WordInterop.Document document;

                var extension = Path.GetExtension(inputFile);

                if (extension != null && extension.ToUpperInvariant() == ".TXT")
                    document = word.Documents.OpenNoRepairDialog(inputFile, false, true, false, "dummypassword",
                        Format: WordInterop.WdOpenFormat.wdOpenFormatUnicodeText,
                        OpenAndRepair: repairMode,
                        NoEncodingDialog: true);
                else
                    document = word.Documents.OpenNoRepairDialog(inputFile, false, true, false, "dummypassword",
                        OpenAndRepair: repairMode,
                        NoEncodingDialog: true);

                // This will lock or unlock all form fields in a Word document so that auto fill 
                // and date/time field do or don't get updated automaticly when converting
                if (document.Fields.Count > 0)
                {
                    foreach (WordInterop.Field field in document.Fields)
                        field.Locked = true;
                }

                return document;
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' seems to be corrupt, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return Open(word, inputFile, true);
            }
        }
        #endregion
    }
}
