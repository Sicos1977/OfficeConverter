using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using WordInterop = Microsoft.Office.Interop.Word;

//
// Word.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2019 Magic-Sessions. (www.magic-sessions.com)
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
    ///     This class is used as a placeholder for all Word related methods
    /// </summary>
    internal class Word : IDisposable
    {
        #region Fields
        /// <summary>
        ///     Word version number
        /// </summary>
        private readonly int _versionNumber;

        /// <summary>
        ///     <see cref="WordInterop.ApplicationClass" />
        /// </summary>
        private WordInterop.ApplicationClass _word;

        /// <summary>
        ///     A <see cref="Process" /> object to Word
        /// </summary>
        private Process _wordProcess;

        /// <summary>
        ///     Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;
        #endregion

        #region Properties
        /// <summary>
        ///     Returns <c>true</c> when Word is running
        /// </summary>
        /// <returns></returns>
        private bool IsWordRunning
        {
            get
            {
                if (_wordProcess == null)
                    return false;

                _wordProcess.Refresh();
                return !_wordProcess.HasExited;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     This constructor checks to see if all requirements for a successful conversion are here.
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the registry could not be read to determine Word version</exception>
        internal Word()
        {
            Logger.WriteToLog("Checking what version of Word is installed");

            try
            {
                var baseKey = Registry.ClassesRoot;
                var subKey = baseKey.OpenSubKey(@"Word.Application\CurVer");
                if (subKey != null)
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // Word 2003
                        case "WORD.APPLICATION.11":
                            _versionNumber = 11;
                            Logger.WriteToLog("Word 2003 is installed");
                            break;

                        // Word 2007
                        case "WORD.APPLICATION.12":
                            _versionNumber = 12;
                            Logger.WriteToLog("Word 2007 is installed");
                            break;

                        // Word 2010
                        case "WORD.APPLICATION.14":
                            _versionNumber = 14;
                            Logger.WriteToLog("Word 2010 is installed");
                            break;

                        // Word 2013
                        case "WORD.APPLICATION.15":
                            _versionNumber = 15;
                            Logger.WriteToLog("Word 2013 is installed");
                            break;

                        // Word 2016
                        case "WORD.APPLICATION.16":
                            _versionNumber = 16;
                            Logger.WriteToLog("Word 2016 is installed");
                            break;

                        // Word 2019
                        case "WORD.APPLICATION.17":
                            _versionNumber = 17;
                            Logger.WriteToLog("Word 2019 is installed");
                            break;

                        default:
                            throw new OCConfiguration("Could not determine Word version");
                    }
                else
                    throw new OCConfiguration("Could not find registry key Word.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check Word version", exception);
            }
        }
        #endregion

        #region StartWord
        /// <summary>
        ///     Starts Word
        /// </summary>
        private void StartWord()
        {
            if (IsWordRunning)
            {
                Logger.WriteToLog($"Word is already running on PID {_wordProcess.Id}... skipped");
                return;
            }

            Logger.WriteToLog("Starting Word");

            _word = new WordInterop.ApplicationClass
            {
                ScreenUpdating = false,
                DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone,
                DisplayDocumentInformationPanel = false,
                DisplayRecentFiles = false,
                DisplayScrollBars = false,
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                Visible = false
            };

            _word.Options.UpdateLinksAtOpen = false;
            _word.Options.ConfirmConversions = false;
            _word.Options.SaveInterval = 0;
            _word.Options.SaveNormalPrompt = false;
            _word.Options.SavePropertiesPrompt = false;
            _word.Options.AllowReadingMode = false;
            _word.Options.WarnBeforeSavingPrintingSendingMarkup = false;
            _word.Options.UpdateFieldsAtPrint = false;
            _word.Options.UpdateLinksAtOpen = false;
            _word.Options.UpdateLinksAtPrint = false;

            var captionGuid = Guid.NewGuid().ToString();
            _word.Caption = captionGuid;

            var processId = ProcessHelpers.GetProcessIdByWindowTitle(captionGuid);

            if (!processId.HasValue)
                throw new OCConfiguration("Could not determine Word process by title");

            _wordProcess = Process.GetProcessById(processId.Value);

            Logger.WriteToLog($"Word started with process id {_wordProcess.Id}");
        }
        #endregion

        #region StopWord
        /// <summary>
        ///     Stops Word
        /// </summary>
        private void StopWord()
        {
            if (IsWordRunning)
            {
                Logger.WriteToLog("Stopping Word");

                try
                {
                    _word.Quit(false);
                }
                catch(Exception exception)
                {
                    Logger.WriteToLog($"Word did not shutdown gracefully, exception: {ExceptionHelpers.GetInnerException(exception)}");
                }

                var counter = 0;

                // Give Word 2 seconds to close
                while (counter < 200)
                {
                    if (!IsWordRunning) break;
                    counter++;
                    Thread.Sleep(10);
                }

                if (IsWordRunning)
                {
                    Logger.WriteToLog($"Word did not shutdown gracefully in 2 seconds ... killing it on process id {_wordProcess.Id}");
                    _wordProcess.Kill();
                    Logger.WriteToLog("Word process killed");
                }
                else
                    Logger.WriteToLog("Word stopped");
            }

            if (_word != null)
            {
                Marshal.ReleaseComObject(_word);
                _word = null;
            }

            _wordProcess = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts a Word document to PDF
        /// </summary>
        /// <param name="inputFile">The Word input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        internal void Convert(string inputFile, string outputFile)
        {
            DeleteResiliencyKeys();

            WordInterop.DocumentClass document = null;

            try
            {
                StartWord();

                document = (WordInterop.DocumentClass) OpenDocument(inputFile, false);

                // Do not remove this line!!
                // This is yet another solution to a weird Office problem. Sometimes there
                // are Word documents with images in it that take some time to load. When
                // we remove the line below the ExportAsFixedFormat method will be called 
                // before the images are loaded thus resulting in an un endless loop somewhere
                // in this method.
                // ReSharper disable once UnusedVariable
                var count = document.ComputeStatistics(WordInterop.WdStatistic.wdStatisticPages);

                Logger.WriteToLog($"Exporting document to PDF file '{outputFile}'");
                document.ExportAsFixedFormat(outputFile, WordInterop.WdExportFormat.wdExportFormatPDF);
                Logger.WriteToLog("Document exported to PDF");
            }
            catch (Exception)
            {
                StopWord();
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
        ///     Opens the <paramref name="inputFile" /> and returns it as an <see cref="WordInterop.Document" /> object
        /// </summary>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile" /> is opened in repair mode</param>
        /// <returns></returns>
        private WordInterop.Document OpenDocument(string inputFile, bool repairMode)
        {
            Logger.WriteToLog($"Opening document '{inputFile}'{(repairMode ? " with repair mode" : string.Empty)}");

            try
            {
                WordInterop.Document document;

                var extension = Path.GetExtension(inputFile);

                if (extension != null && extension.ToUpperInvariant() == ".TXT")
                    document = _word.Documents.OpenNoRepairDialog(inputFile, false, true, false, "dummy password",
                        Format: WordInterop.WdOpenFormat.wdOpenFormatUnicodeText,
                        OpenAndRepair: repairMode,
                        NoEncodingDialog: true);
                else
                    document = _word.Documents.OpenNoRepairDialog(inputFile, false, true, false, "dummy password",
                        OpenAndRepair: repairMode,
                        NoEncodingDialog: true);

                // This will lock or unlock all form fields in a Word document so that auto fill 
                // and date/time field do or don't get updated automatic when converting
                if (document.Fields.Count > 0)
                {
                    Logger.WriteToLog("Locking all form fields against modifications");
                    foreach (WordInterop.Field field in document.Fields)
                        field.Locked = true;
                }

                Logger.WriteToLog("Document opened");
                return document;
            }
            catch (Exception exception)
            {
                Logger.WriteToLog(
                    $"ERROR: Failed to open document, exception: '{ExceptionHelpers.GetInnerException(exception)}'");

                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' seems to be corrupt, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return OpenDocument(inputFile, true);
            }
        }
        #endregion

        #region CloseDocument
        /// <summary>
        ///     Closes the opened document and releases any allocated resources
        /// </summary>
        /// <param name="document">The Word document</param>
        private void CloseDocument(WordInterop.Document document)
        {
            if (document == null) return;
            Logger.WriteToLog("Closing document");
            document.Saved = true;
            document.Close(false);
            Marshal.ReleaseComObject(document);
            Logger.WriteToLog("Document closed");
        }
        #endregion

        #region DeleteResiliencyKeys
        /// <summary>
        ///     This method will delete the automatic created Resiliency key. Word uses this registry key
        ///     to make entries to corrupted documents. If there are to many entries under this key Word will
        ///     get slower and slower to start. To prevent this we just delete this key when it exists
        /// </summary>
        private void DeleteResiliencyKeys()
        {
            Logger.WriteToLog("Deleting Word resiliency keys from the registry");

            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Resiliency\DocumentRecovery
                var key = $@"Software\Microsoft\Office\{_versionNumber}.0\Word\Resiliency";

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
        ///     Disposes the running <see cref="_word" />
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            StopWord();
        }
        #endregion
    }
}