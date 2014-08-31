using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
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

namespace OfficeConverter
{
    /// <summary>
    /// With this class an Microsoft Office document can be converted to PDF format. Microsoft Office 2007 
    /// (with PDF export plugin) or higher is needed.
    /// </summary>
    public class Converter
    {
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
                case ".ODP":
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
                case ".ODS":
                    ConvertExcelDocument(inputFile, outputFile);
                    break;

                case ".POT":
                case ".PPT":
                case ".PPS":
                case ".POTM":
                case ".POTX":
                case ".PPSM":
                case ".PPSX":
                case ".PPTM":
                case ".PPTX":
                case ".ODT":
                    ConvertPowerPointPresentation(inputFile, outputFile);
                    break;

                default:
                    throw new OCFileTypeNotSupported("The file '" + Path.GetFileName(inputFile) +
                                                     "' is not supported, only .DOC, .DOCM, .DOCX, .DOT, .DOTM, .ODP, .XLS, .XLSB, .XLSM, .XLSX, .XLT, " +
                                                     ".XLTM, .XLTX, .XLW, .ODS, .POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM, .PPTX and .ODT are supported");
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
        private static void ConvertWordDocument(string inputFile, string outputFile)
        {
            Word.ApplicationClass word = null;
            Word.DocumentClass document = null;

            try
            {
                word = new Word.ApplicationClass
                {
                    ScreenUpdating = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
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
        /// Returns the maximum rows Excel supports
        /// </summary>
        /// <returns></returns>
        private static int GetExcelMaxRows()
        {
            const int excelMaxRowsFrom2003AndBelow = 65535;
            const int excelMaxRowsFrom2007AndUp = 1048576;

            var baseKey = Registry.ClassesRoot;
            var subKey = baseKey.OpenSubKey(@"Excel.Application\CurVer");
            if (subKey != null)
            {
                switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                {
                    case "EXCEL.APPLICATION.11":
                        return excelMaxRowsFrom2003AndBelow;

                    case "EXCEL.APPLICATION.12":
                    case "EXCEL.APPLICATION.14":
                    case "EXCEL.APPLICATION.15":
                        return excelMaxRowsFrom2007AndUp;
                }
            }

            throw new Exception("Could not read registry to check Excel version");
        }

        /// <summary>
        /// Converts a Excel document to PDF
        /// </summary>
        /// <param name="inputFile">The Excel input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private static void ConvertExcelDocument(string inputFile, string outputFile)
        {
            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            string tempFileName = null;

            try
            {
                excel = new Excel.ApplicationClass
                {
                    //ScreenUpdating = false,
                    DisplayAlerts = false,
                    DisplayDocumentInformationPanel = false,
                    DisplayRecentFiles = false,
                    DisplayScrollBars = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                };

                // TODO: Set specific culture 
                //excel.DecimalSeparator = ci.NumberFormat.NumberDecimalSeparator;
                //excel.ThousandsSeparator = ci.NumberFormat.NumberGroupSeparator;


                var extension = Path.GetExtension(inputFile);
                if (string.IsNullOrWhiteSpace(extension))
                    extension = string.Empty;

                if (extension.ToUpperInvariant() == ".CSV")
                {
                    // Yes this look somewhat weird but we have to change the extension if we want to handle
                    // CSV files with different kind of separators. Otherwhise Excel will always overrule whatever
                    // setting we make to open a file
                    tempFileName = Path.GetTempFileName() + Guid.NewGuid() + ".txt";
                    File.Copy(inputFile, tempFileName);
                }

                workbook = OpenExcelWorkbook(excel, inputFile, extension, false);
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

                if (!string.IsNullOrEmpty(tempFileName) && File.Exists(tempFileName))
                    File.Delete(tempFileName);
            }
        }
        #endregion

        #region OpenExcelWorkbook
        /// <summary>
        /// Returns the seperator and textqualifier that is used in the CSV file
        /// </summary>
        /// <param name="inputFile">The inputfile</param>
        /// <param name="separator">The separator that is used</param>
        /// <param name="textQualifier">The text qualifier</param>
        /// <returns></returns>
        private static void GetCsvSeperator(string inputFile, out string separator, out Excel.XlTextQualifier textQualifier)
        {
            separator = string.Empty;
            textQualifier = Excel.XlTextQualifier.xlTextQualifierNone;
            
            using (var streamReader = new StreamReader(inputFile))
            {
                var line = string.Empty;
                while (string.IsNullOrEmpty(line))
                    line = streamReader.ReadLine();

                if (line.Contains(";")) separator = ";";
                else if (line.Contains(",")) separator = ",";
                else if (line.Contains("\t")) separator = "\t";
                else if (line.Contains(" ")) separator = " ";

                if (line.Contains("\"")) textQualifier = Excel.XlTextQualifier.xlTextQualifierDoubleQuote;
                else if (line.Contains(",")) textQualifier = Excel.XlTextQualifier.xlTextQualifierSingleQuote;
            }
        }

        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="Excel.Workbook"/> object
        /// </summary>
        /// <param name="excel">The <see cref="Excel.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="extension">The file extension</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private static Excel.Workbook OpenExcelWorkbook(Excel._Application excel,
                                                        string inputFile,
                                                        string extension,
                                                        bool repairMode)
        {
            try
            {
                switch (extension.ToUpperInvariant())
                {
                    case ".CSV":

                        var count = File.ReadLines(inputFile).Count();
                        var excelMaxRows = GetExcelMaxRows();
                        if (count > excelMaxRows)
                            throw new OCCsvFileLimitExceeded("The input CSV file has more then " + excelMaxRows +
                                                             " rows, the installed Excel version supports only " +
                                                             excelMaxRows + " rows");

                        string separator;
                        Excel.XlTextQualifier textQualifier;

                        GetCsvSeperator(inputFile, out separator, out textQualifier);

                        switch (separator)
                        {
                            case ";":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited,
                                    textQualifier, true, false, true);
                                return excel.ActiveWorkbook;

                            case ",":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, true);
                                return excel.ActiveWorkbook;

                            case "\t":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, true);
                                return excel.ActiveWorkbook;

                            case " ":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, false, true);
                                return excel.ActiveWorkbook;

                            default:
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    Excel.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, true);
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

                return OpenExcelWorkbook(excel, inputFile, extension, true);
            }

        }
        #endregion

        #region ConvertPowerPointPresentation
        /// <summary>
        /// Converts a PowerPoint document to PDF
        /// </summary>
        /// <param name="inputFile">The PowerPoint input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        private static void ConvertPowerPointPresentation(string inputFile, string outputFile)
        {
            PowerPoint.ApplicationClass powerPoint = null;
            PowerPoint.Presentation presentation = null;

            try
            {
                powerPoint = new PowerPoint.ApplicationClass
                {
                    DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone,
                    DisplayDocumentInformationPanel = false,
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                };

                presentation = powerPoint.Presentations.Open(inputFile, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                presentation.ExportAsFixedFormat(outputFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
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
    }
}
