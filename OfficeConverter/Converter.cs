using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

/*
   Copyright 2013-2014 Kees van Spelde

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

//http://omegacoder.com/?p=555
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;

namespace OfficeConverter
{
    public class Converter
    {
        #region CheckFileNameAndOutputFolder
        /// <summary>
        /// Checks if the <see cref="inputFile"/> and the path to the <see cref="outputFile"/> is valid
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        /// <exception cref="ArgumentNullException">Raised when the <see cref="inputFile"/> or <see cref="outputFile"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <see cref="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the path to the <see cref="outputFile"/> does not exists</exception>
        private static void CheckFileNameAndOutputFolder(string inputFile, string outputFile)
        {
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentNullException(inputFile);

            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentNullException(outputFile);

            if (!File.Exists(inputFile))
                throw new FileNotFoundException("Could not find the input file '" + inputFile + "'");

            var directoryInfo = new FileInfo(outputFile).Directory;
            if (directoryInfo == null) return;
            var outputFolder = directoryInfo.FullName;

            if (!Directory.Exists(outputFolder))
                throw new DirectoryNotFoundException("The output folder '" + outputFolder + "' does not exist");
        }
        #endregion

        #region ConvertToFolder
        /// <summary>
        /// Converts the given <see cref="inputFile"/> to PDF and saves it as the <see cref="outputFile"/>
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFile">The output folder</param>
        /// <returns>List with files or en empty list when there are nog embedded files</returns>
        /// <exception cref="ArgumentNullException">Raised when the <see cref="inputFile"/> or <see cref="outputFile"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <see cref="inputFile"/> does not exist</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the path to the <see cref="outputFile"/> does not exists</exception>
        /// <exception cref="OCFileIsCorrupt">Raised when the <see cref="inputFile" /> is corrupt</exception>
        /// <exception cref="OCFileTypeNotSupported">Raised when the <see cref="inputFile"/> is not supported</exception>
        /// <exception cref="OCFileIsPasswordProtected">Raised when the <see cref="inputFile"/> is password protected</exception>
        public void ConvertToFolder(string inputFile, string outputFile)
        {
            CheckFileNameAndOutputFolder(inputFile, outputFile);

            var extension = Path.GetExtension(inputFile);
            if (extension != null)
                extension = extension.ToUpperInvariant();

            switch (extension)
            {
                //case ".ODT":
                //case ".ODS":
                //case ".ODP":
                //    return ExtractFromOpenDocumentFormat(inputFile, outputFolder);

                case ".DOC":
                case ".DOCM":
                case ".DOCX":
                case ".DOT":
                case ".DOTM":
                    ConvertWord(inputFile, outputFile);
                    break;

                case ".XLS":
                case ".XLT":
                case ".XLW":
                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                    ConvertExcel(inputFile, outputFile);
                    break;

                case ".POT":
                case ".PPT":
                case ".PPS":
                    // PowerPoint 97 - 2003
                    //return ExtractFromPowerPointBinaryFormat(inputFile, outputFolder);

                case ".POTM":
                case ".POTX":
                case ".PPSM":
                case ".PPSX":
                case ".PPTM":
                case ".PPTX":
                    // PowerPoint 2007 - 2013
                    //return ExtractFromOfficeOpenXmlFormat(inputFile, "/ppt/embeddings/", outputFolder);

                default:
                    throw new OCFileTypeNotSupported("The file '" + Path.GetFileName(inputFile) +
                                                     "' is not supported, only .DOC, .DOCM, .DOCX, .DOT, .DOTM, .XLS, .XLSB, .XLSM, .XLSX, .XLT, " +
                                                     ".XLTM, .XLTX, .XLW, .POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM and .PPTX are supported");
            }
        }
        #endregion

        #region Word
        /// <summary>
        /// This method will open a Word document and save it as PDF
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFolder"></param>
        /// <returns>The path and filename to the saved PDF</returns>
        /// <exception cref="OCFileIsPasswordProtected">Raised when the <see cref="inputFile"/> is password protected</exception>
        /// <exception cref="OCFileIsCorrupt">Raised when the <see cref="inputFile"/> is corrupt</exception>
        public void ConvertWord(string inputFile, string outputFolder)
        {
            Word.Application word = null;
            Word.Document document = null;

            try
            {
                word = new Word.Application();
                word.ScreenUpdating = false;
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
                word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                word.DisplayDocumentInformationPanel = false;
                word.DisplayRecentFiles = false;
                word.DisplayScrollBars = false;
                word.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                // Open the Word document
                document = OpenWordDocument(word, inputFile, false);
                
                word.DisplayAutoCompleteTips = false;
                word.DisplayScreenTips = false;
                word.DisplayStatusBar = false;

                //var outputFile = outputFolder + Path.GetTempFileName()

                // This will lock all form fields in a Word document so that auto fill 
                // and date/time field do or don't get updated automaticly when saving as PDF
                if (document.Fields.Count > 0)
                {
                    foreach (Word.Field field in document.Fields)
                        field.Locked = true;
                }

                document.ExportAsFixedFormat(outputFolder, 
                                             ExportFormat: Word.WdExportFormat.wdExportFormatPDF, 
                                             OpenAfterExport: false,
                                             OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint);

            }
            catch (COMException comException)
            {
                // TODO: Log exceptie
                throw;
            }
            finally
            {
                if (document != null)
                {
                    document.Saved = true;
                    document.Close(SaveChanges: false);
                    Marshal.ReleaseComObject(document);
                }

                if (word != null)
                {
                    word.Quit(SaveChanges: false);
                    Marshal.ReleaseComObject(word);
                }
            }
        }

        /// <summary>
        /// This method will open the given <see cref="inputFile"/> and return a <see cref="Word.Document"/> object
        /// </summary>
        /// <param name="word">The <see cref="Word.Application"/> object</param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode"></param>
        /// <returns></returns>
        /// <exception cref="OCFileIsPasswordProtected">Raised when the <see cref="inputFile"/> is password protected</exception>
        /// <exception cref="OCFileIsCorrupt">Raised when the <see cref="inputFile"/> is corrupt</exception>
        public Word.Document OpenWordDocument(Word.Application word,
                                              string inputFile,
                                              bool repairMode)
        {
            Word.Document document;

            try
            {
                var extension = Path.GetExtension(inputFile);
                if (!string.IsNullOrEmpty(extension))
                    extension = extension.ToUpperInvariant();

                switch (extension)
                {
                    case ".TXT":
                        document = word.Documents.Open(inputFile, 
                                                     ConfirmConversions: false, 
                                                     ReadOnly: true, 
                                                     AddToRecentFiles: false, 
                                                     PasswordDocument: "dummypassword", 
                                                     Format: Word.WdOpenFormat.wdOpenFormatUnicodeText,
                                                     OpenAndRepair: true,
                                                     NoEncodingDialog: true);
                        break;

                    default:
                        document = word.Documents.Open(inputFile,
                                                     ConfirmConversions: false,
                                                     ReadOnly: true,
                                                     AddToRecentFiles: false,
                                                     PasswordDocument: "dummypassword",
                                                     OpenAndRepair: true,
                                                     NoEncodingDialog: true);
                        break;
                }
            }
            catch (COMException comException)
            {
                // 5408 = password protected
                if (comException.ErrorCode == 5408)
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' is corrupt");

                document = OpenWordDocument(word, inputFile, true);
            }

            return document;
        }
        #endregion

        #region Excel

        public string ConvertExcel(string inputFile, string outputFile)
        {
            throw new NotImplementedException();
        }

        public Word.Document OpenExcelWorkbook(Excel.Application excel,
                                               string inputFile,
                                               bool repairMode)
        {
            Excel.Workbook workbook;

            try
            {
                var extension = Path.GetExtension(inputFile);
                if (!string.IsNullOrEmpty(extension))
                    extension = extension.ToUpperInvariant();

                switch (extension)
                {
                    case ".TXT":
                        workbook = word.Documents.Open(inputFile,
                                                     ConfirmConversions: false,
                                                     ReadOnly: true,
                                                     AddToRecentFiles: false,
                                                     PasswordDocument: "dummypassword",
                                                     Format: Word.WdOpenFormat.wdOpenFormatUnicodeText,
                                                     OpenAndRepair: true,
                                                     NoEncodingDialog: true);
                        break;

                    default:
                        workbook = word.Documents.Open(inputFile,
                                                     ConfirmConversions: false,
                                                     ReadOnly: true,
                                                     AddToRecentFiles: false,
                                                     PasswordDocument: "dummypassword",
                                                     OpenAndRepair: true,
                                                     NoEncodingDialog: true);
                        break;
                }
            }
            catch (COMException comException)
            {
                // 5408 = password protected
                if (comException.ErrorCode == 5408)
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' is corrupt");

                workbook = OpenWordDocument(word, inputFile, true);
            }

            return workbook;
        }
        #endregion

    }
}
