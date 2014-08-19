using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    /// <summary>
    /// With this class an Microsoft Office document can be converted to PDF format. Microsoft Office 2007 
    /// (with PDF export plugin) or higher is needed.
    /// </summary>
    public class Converter
    {
        #region Consts
        /// <summary>
        /// Maximum rows on an Excel 2007 or higher worksheet
        /// </summary>
        const int ExcelMaxRows = 1048576;

        /// <summary>
        /// Maximum columns on an Excel 2007 or higher worksheet
        /// </summary>
        const int ExcelMaxColumns = 16384;
        #endregion

        #region CheckFileNameAndOutputFolder
        /// <summary>
        /// Checks if the <paramref name="inputFile"/> and the folder where the <paramref name="outputFile"/> is written exists
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        /// <exception cref="ArgumentNullException">Raised when the <paramref name="inputFile"/> or <paramref name="outputFile"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <paramref name="inputFile"/> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the folder where the <paramref name="outputFile"/> is written does not exists</exception>
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

        #region Convert
        /// <summary>
        /// Converts the <paramref name="inputFile"/> to PDF and saves it as the <paramref name="outputFile"/>
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFile">The output file with full path</param>
        /// <exception cref="ArgumentNullException">Raised when the <paramref name="inputFile"/> or <paramref name="outputFile"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <paramref name="inputFile"/> does not exist</exception>
        /// <exception cref="DirectoryNotFoundException">Raised when the folder where the <paramref name="outputFile"/> is written does not exists</exception>
        /// <exception cref="OCFileIsCorrupt">Raised when the <paramref name="inputFile" /> is corrupt</exception>
        /// <exception cref="OCFileTypeNotSupported">Raised when the <paramref name="inputFile"/> is not supported</exception>
        /// <exception cref="OCFileIsPasswordProtected">Raised when the <paramref name="inputFile"/> is password protected</exception>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        public void Convert(string inputFile, string outputFile)
        {
            CheckFileNameAndOutputFolder(inputFile, outputFile);

            var extension = Path.GetExtension(inputFile);
            if (extension != null)
                extension = extension.ToUpperInvariant();

            switch (extension)
            {
                case ".DOC":
                case ".DOT":
                case ".DOCM":
                case ".DOCX":
                case ".DOTM":
                    ConvertWordDocument(inputFile, outputFile);
                    break;

                case ".XLS":
                case ".XLT":
                case ".XLW":
                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                case ".CSV":
                    ConvertExcelDocument(inputFile, outputFile);
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

        #region ConvertWordDocument
        /// <summary>
        /// Converts a Word document to PDF
        /// </summary>
        /// <param name="inputFile">The Word input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        private void ConvertWordDocument(string inputFile, string outputFile)
        {
            Word.ApplicationClass word = null;
            Word.DocumentClass document = null;

            try
            {
                word = new Word.ApplicationClass()
                {
                    ScreenUpdating = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    //AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
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

                document = (Word.DocumentClass) OpenWordDocument(word, inputFile, false);
                
                word.DisplayAutoCompleteTips = false;
                word.DisplayScreenTips = false;
                word.DisplayStatusBar = false;

                document.ExportAsFixedFormat(outputFile, Word.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                if (document != null)
                {
                    document.Saved = true;
                    document.Close();
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

        #region OpenWordDocument
        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="Word.Document"/> object
        /// </summary>
        /// <param name="word">The <see cref="Word.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        private static Word.Document OpenWordDocument(Word._Application word,
                                                           string inputFile,
                                                           bool repairMode)
        {
            try
            {
                Word.Document document;

                var extension = Path.GetExtension(inputFile);

                if (extension != null && extension.ToUpperInvariant() == ".TXT")
                    document = word.Documents.OpenNoRepairDialog(inputFile, false, true, false, "dummypassword",
                        Format: Word.WdOpenFormat.wdOpenFormatUnicodeText,
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
                    foreach (Word.Field field in document.Fields)
                        field.Locked = true;
                }

                return document;
            }
            catch (COMException comException)
            {
                if (comException.ErrorCode == 5408)
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) + "' is password protected");

                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' seems to be corrupt");

                return OpenWordDocument(word, inputFile, true);
            }
        }
        #endregion

        #region ConvertExcelDocument
        /// <summary>
        /// Converts a Excel document to PDF
        /// </summary>
        /// <param name="inputFile">The Excel input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private void ConvertExcelDocument(string inputFile, string outputFile)
        {
            Excel.Application excel = null;
            Excel.Workbook workbook = null;

            try
            {
                excel = new Excel.ApplicationClass
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                };

                workbook = OpenExcelWorkbook(excel, inputFile, false);
                workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Saved = true;
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }
            }
        }
        #endregion

        #region OpenExcelWorkbook
        /// <summary>
        /// Returns the seperator that is used in the CSV file. If no seperator is found an empty string is returned
        /// </summary>
        /// <param name="inputFile"></param>
        /// <returns></returns>
        private static string GetCsvSeperator(string inputFile)
        {
            using (var streamReader = new StreamReader(inputFile))
            {
                var line = string.Empty;
                while (string.IsNullOrEmpty(line))
                    line = streamReader.ReadLine();

                if (line.Contains(";")) return ";";
                if (line.Contains(",")) return ",";
                if (line.Contains("\t")) return "\t";
                if (line.Contains(" ")) return " ";
            }

            return string.Empty;
        }

        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="Excel.Workbook"/> object
        /// </summary>
        /// <param name="excel">The <see cref="Excel.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private Excel.Workbook OpenExcelWorkbook(Excel._Application excel,
                                                 string inputFile,
                                                 bool repairMode)
        {
            try
            {

                var extension = Path.GetExtension(inputFile);
                if (string.IsNullOrWhiteSpace(extension))
                    extension = string.Empty;

                var count = File.ReadLines(inputFile).Count();
                if (count > ExcelMaxRows)
                    throw new OCCsvFileLimitExceeded("The input CSV file has more then " + ExcelMaxRows + " rows");

                switch (extension.ToUpperInvariant())
                {
                    case ".CSV":

                        var seperator = GetCsvSeperator(inputFile);

                        switch (seperator)
                        {
                            case ";":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing, Excel.XlTextParsingType.xlDelimited,
                                    Excel.XlTextQualifier.xlTextQualifierNone, Type.Missing, false, true, false, false);
                                return excel.ActiveWorkbook;

                            case ",":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone,
                                    Type.Missing, false, false, true, false);
                                return excel.ActiveWorkbook;

                            case "\t":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone,
                                    Type.Missing, true, false, false, false);
                                return excel.ActiveWorkbook;

                            case " ":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone,
                                    Type.Missing, false, false, false, true);
                                return excel.ActiveWorkbook;

                            default:
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, Excel.XlTextQualifier.xlTextQualifierNone,
                                    Type.Missing, false, true, false, false);
                                return excel.ActiveWorkbook;
                        }

                    default:

                        if (repairMode)
                            return excel.Workbooks.Open(inputFile, false, true,
                                Password: "dummypassword",
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false,
                                CorruptLoad: Excel.XlCorruptLoad.xlRepairFile);

                        return excel.Workbooks.Open(inputFile, false, true,
                            IgnoreReadOnlyRecommended: true,
                            AddToMru: false);

                }
            }
            catch (COMException comException)
            {
                if (comException.ErrorCode == 5408)
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) + "' is password protected");

                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) + "' seems to be corrupt");

                return OpenExcelWorkbook(excel, inputFile, true);
            }

        }
        #endregion
    }
}
