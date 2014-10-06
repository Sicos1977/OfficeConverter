using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CompoundFileStorage;
using CompoundFileStorage.Exceptions;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

/*
   Copyright 2014 Kees van Spelde

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
    #region Interface IReader
    /// <summary>
    /// Interface to make Reader class COM exposable
    /// </summary>
    public interface IConverter
    {
        /// <summary>
        /// Converts the <paramref name="inputFile"/> to PDF and saves it as the <paramref name="outputFile"/>
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFile">The output file with full path</param>
        /// <returns>Returns true when the conversion is succesfull, false is retournerd when an exception occurred. 
        /// The exception can be retrieved with the <see cref="GetErrorMessage"/> method</returns>
        [DispId(1)]
        bool ConvertFromCom(string inputFile, string outputFile);

        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        [DispId(2)]
        string GetErrorMessage();
    }
    #endregion

    /// <summary>
    /// With this class an Microsoft Office document can be converted to PDF format. Microsoft Office 2007 
    /// (with PDF export plugin) or higher is needed.
    /// </summary>
    [Guid("4F474ED1-70C5-47D4-8EEF-CDB3E1149455")]
    [ComVisible(true)]
    public class Converter : IConverter
    {
        #region Fields
        /// <summary>
        /// Contains an error message when something goes wrong in the <see cref="ConvertFromCom"/> method.
        /// This message can be retreived with the GetErrorMessage. This way we keep .NET exceptions inside
        /// when this code is called from a COM language
        /// </summary>
        private string _errorMessage;
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
        /// <returns>Returns true when the conversion is succesfull, false is retournerd when an exception occurred. 
        /// The exception can be retrieved with the <see cref="GetErrorMessage"/> method</returns>
        public bool ConvertFromCom(string inputFile, string outputFile)
        {
            try
            {
                _errorMessage = string.Empty;
                Convert(inputFile, outputFile);
                return true;
            }
            catch (Exception exception)
            {
                _errorMessage = ExceptionHelpers.GetInnerException(exception);
                return false;
            }
        }

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
                    if (WordFileIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    ConvertWithWord(inputFile, outputFile);
                    break;

                case ".ODT":
                    if (OpenDocumentFormatIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    ConvertWithWord(inputFile, outputFile);
                    break;

                case ".RTF":
                    ConvertWithWord(inputFile, outputFile);
                    break;

                case ".XLS":
                case ".XLT":
                case ".XLW":
                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                    if (ExcelFileIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected"); 

                    ConvertWithExcel(inputFile, outputFile);
                    break;
                
                case ".CSV":
                    ConvertWithExcel(inputFile, outputFile);
                    break;

                case ".ODS":
                    if (OpenDocumentFormatIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    ConvertWithExcel(inputFile, outputFile);
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
                    if (PowerPointFileIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) + "' is password protected"); 
            
                    ConvertWithPowerPoint(inputFile, outputFile);
                    break;

                case ".ODP":
                    if (OpenDocumentFormatIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    ConvertWithPowerPoint(inputFile, outputFile);
                    break;

                default:
                    throw new OCFileTypeNotSupported("The file '" + Path.GetFileName(inputFile) +
                                                     "' is not supported, only .DOC, .DOCM, .DOCX, .DOT, .DOTM, .RTF, .ODT, .XLS, .XLSB, .XLSM, .XLSX, .XLT, " +
                                                     ".XLTM, .XLTX, .XLW, .ODS, .POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM, .PPTX and .ODP are supported");
            }
        }
        #endregion

        #region GetErrorMessage
        /// <summary>
        /// Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        public string GetErrorMessage()
        {
            return _errorMessage;
        }
        #endregion

        #region ExtractFromOpenDocumentFormat
        /// <summary>
        /// Returns true when the <paramref name="inputFile"/> is password protected
        /// </summary>
        /// <param name="inputFile">The OpenDocument format file</param>
        public bool OpenDocumentFormatIsPasswordProtected(string inputFile)
        {
            var zipFile = new ZipFile(inputFile);

            // Check if the file is password protected
            var manifestEntry = zipFile.FindEntry("META-INF/manifest.xml", true);
            if (manifestEntry != -1)
            {
                using (var manifestEntryStream = zipFile.GetInputStream(manifestEntry))
                using (var manifestEntryMemoryStream = new MemoryStream())
                {
                    manifestEntryStream.CopyTo(manifestEntryMemoryStream);
                    manifestEntryMemoryStream.Position = 0;
                    using (var streamReader = new StreamReader(manifestEntryMemoryStream))
                    {
                        var manifest = streamReader.ReadToEnd();
                        if (manifest.ToUpperInvariant().Contains("ENCRYPTION-DATA"))
                            return true;
                    }
                }
            }

            return false;
        }
        #endregion

        #region ConvertWithWord
        /// <summary>
        /// Converts a Word document to PDF
        /// </summary>
        /// <param name="inputFile">The Word input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        private static void ConvertWithWord(string inputFile, string outputFile)
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

                document = (Word.DocumentClass) OpenWordFile(word, inputFile, false);
                
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

        #region WordFileIsPasswordProtected
        /// <summary>
        /// Returns true when the Word file is password protected
        /// </summary>
        /// <param name="inputFile">The Word file to check</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">Raised when the file is corrupt</exception>
        public static bool WordFileIsPasswordProtected(string inputFile)
        {
            try
            {
                using (var compoundFile = new CompoundFile(inputFile))
                {
                    if (compoundFile.RootStorage.ExistsStream("EncryptedPackage")) return true; 
                    if (!compoundFile.RootStorage.ExistsStream("WordDocument"))
                        throw new OCFileIsCorrupt("Could not find the WordDocument stream in the file '" + compoundFile.FileName + "'"); 
                    
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
            catch (CFFileFormatException)
            {
                // It seems the file is just a normal Microsoft Office 2007 and up Open XML file
                return false;
            }
        }
        #endregion

        #region OpenWordFile
        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="Word.Document"/> object
        /// </summary>
        /// <param name="word">The <see cref="Word.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        private static Word.Document OpenWordFile(Word._Application word,
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
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' seems to be corrupt, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return OpenWordFile(word, inputFile, true);
            }
        }
        #endregion

        #region ConvertWithExcel
        /// <summary>
        /// If you want to run this code on a server the following folders must exists, if they don't
        /// then you can't use Excel to convert files to PDF
        /// </summary>
        private static void CheckIfSystemProfileDesktopDirectoryExists()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                var x64DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"SysWOW64\config\systemprofile\desktop");
                if (!Directory.Exists(x64DesktopPath))
                {
                    try
                    {
                        Directory.CreateDirectory(x64DesktopPath);
                    }
                    catch (Exception exception)
                    {
                        throw new IOException("Can't create directory '" + x64DesktopPath + "', error: " +
                                              ExceptionHelpers.GetInnerException(exception));
                    }
                }
            }

            var x86DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                @"System32\config\systemprofile\desktop");
            if (!Directory.Exists(x86DesktopPath))
            {
                try
                {
                    Directory.CreateDirectory(x86DesktopPath);
                }
                catch (Exception exception)
                {
                    throw new IOException("Can't create directory '" + x86DesktopPath + "', error: " +
                                          ExceptionHelpers.GetInnerException(exception));
                }
            }
        }

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
        /// Converts a Excel sheet to PDF
        /// </summary>
        /// <param name="inputFile">The Excel input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private static void ConvertWithExcel(string inputFile, string outputFile)
        {
            CheckIfSystemProfileDesktopDirectoryExists();

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
                    inputFile = tempFileName;
                }

                workbook = OpenExcelFile(excel, inputFile, extension, false);
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

        #region ExcelFileIsPasswordProtected
        /// <summary>
        /// Returns true when the Excel file is password protected
        /// </summary>
        /// <param name="inputFile">The Excel file to check</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">Raised when the file is corrupt</exception>
        public static bool ExcelFileIsPasswordProtected(string inputFile)
        {
            try
            {
                using (var compoundFile = new CompoundFile(inputFile))
                {
                    if (compoundFile.RootStorage.ExistsStream("EncryptedPackage")) return true;
                    if (!compoundFile.RootStorage.ExistsStream("WorkBook"))
                        throw new OCFileIsCorrupt("Could not find the WorkBook stream in the file '" +
                                                  compoundFile.FileName + "'");

                    var stream = compoundFile.RootStorage.GetStream("WorkBook") as CFStream;
                    if (stream == null) return false;

                    var bytes = stream.GetData();
                    using (var memoryStream = new MemoryStream(bytes))
                    using (var binaryReader = new BinaryReader(memoryStream))
                    {
                        // Get the record type, at the beginning of the stream this should always be the BOF
                        var recordType = binaryReader.ReadUInt16();

                        // Something seems to be wrong, we would expect a BOF but for some reason it isn't so stop it
                        if (recordType != 0x809)
                            throw new OCFileIsCorrupt("The file '" + Path.GetFileName(compoundFile.FileName) +
                                                      "' is corrupt");

                        var recordLength = binaryReader.ReadUInt16();
                        binaryReader.BaseStream.Position += recordLength;

                        // Search after the BOF for the FilePass record, this starts with 2F hex
                        recordType = binaryReader.ReadUInt16();
                        return (recordType == 0x2F);
                    }
                }
            }
            catch (CFFileFormatException)
            {
                // It seems the file is just a normal Microsoft Office 2007 and up Open XML file
                return false;
            }
        }
        #endregion

        #region OpenExcelFile
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
        private static Excel.Workbook OpenExcelFile(Excel._Application excel,
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
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' seems to be corrupt, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return OpenExcelFile(excel, inputFile, extension, true);
            }
        }
        #endregion

        #region ConvertWithPowerPoint
        /// <summary>
        /// Converts a PowerPoint document to PDF
        /// </summary>
        /// <param name="inputFile">The PowerPoint input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        private static void ConvertWithPowerPoint(string inputFile, string outputFile)
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
                
                presentation = OpenPowerPointFile(powerPoint, inputFile, false);
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

        #region PowerPointFileIsPasswordProtected
        /// <summary>
        /// Returns true when the binary PowerPoint file is password protected
        /// </summary>
        /// <param name="inputFile">The PowerPoint file to check</param>
        /// <returns></returns>
        private static bool PowerPointFileIsPasswordProtected(string inputFile)
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
            catch (CFFileFormatException)
            {
                // It seems the file is just a normal Microsoft Office 2007 and up Open XML file
                return false;
            }
        }
        #endregion

        #region OpenPowerPointFile
        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="PowerPoint.Presentation"/> object
        /// </summary>
        /// <param name="powerPoint">The <see cref="PowerPoint.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">Raised when the <paramref name="inputFile"/> is corrupt and can't be opened in repair mode</exception>
        private static PowerPoint.Presentation OpenPowerPointFile(PowerPoint._Application powerPoint,
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

                return OpenPowerPointFile(powerPoint, inputFile, true);
            }
        }
        #endregion
    }
}
