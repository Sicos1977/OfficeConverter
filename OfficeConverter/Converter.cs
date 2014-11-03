using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CompoundFileStorage;
using CompoundFileStorage.Exceptions;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Biff8;
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
        /// <exception cref="OCFileContainsNoData">Raised when the Microsoft Office file contains no actual data</exception>
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
                case ".MHT":
                case ".WPS":
                case ".WRI":
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
                                                     "' is not supported only, " + Environment.NewLine +
                                                     ".DOC, .DOCM, .DOCX, .DOT, .DOTM, .RTF, .MHT, .WPS, .WRI, .ODT, " + Environment.NewLine +
                                                     ".XLS, .XLSB, .XLSM, .XLSX, .XLT, .XLTM, .XLTX, .XLW, .ODS, " + Environment.NewLine +
                                                     ".POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM, .PPTX and .ODP " + Environment.NewLine +
                                                     " are supported");
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

                // Do not remove this line!!
                // This is yet another solution to a weird Office problem. Sometimes there
                // are Word documents with images in it that take some time to load. When
                // we remove the line below the ExportAsFixedFormat method will be called 
                // before the images are loaded thus resulting in an unendless loop somewhere
                // in this method.
                // ReSharper disable once UnusedVariable
                var count = document.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                
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

        #region CheckIfSystemProfileDesktopDirectoryExists
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
                        throw new OCExcelConfiguration("Can't create directory '" + x64DesktopPath +
                                                       "' Excel needs this folder to work on a server, error: " +
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
                    throw new OCExcelConfiguration("Can't create directory '" + x86DesktopPath +
                                                   "' Excel needs this folder to work on a server, error: " +
                                                   ExceptionHelpers.GetInnerException(exception));
                }
            }
        }
        #endregion

        #region CheckIfPrinterIsInstalled
        /// <summary>
        /// Excel needs a default printer to export to PDF, this method will check if there is one
        /// </summary>
        private static void CheckIfPrinterIsInstalled()
        {
            var result = false;
            foreach (string printerName in PrinterSettings.InstalledPrinters)
            {
                // Retrieve the printer settings.
                var printer = new PrinterSettings { PrinterName = printerName };

                // Check that this is a valid printer.
                // (This step might be required if you read the printer name
                // from a user-supplied value or a registry or configuration file
                // setting.)
                if (printer.IsValid)
                {
                    result = true;
                    break;
                }
            }

            if (!result)
                throw new OCExcelConfiguration("There is no default printer installed, Excel needs one to export to PDF");
        }
        #endregion

        #region ExcelVersionNumber
        /// <summary>
        /// Returns the Excel version number as an integer
        /// </summary>
        /// <exception cref="OCExcelConfiguration">Raised when Excel version could not be determined or the registry could not be read</exception>
        private static int ExcelVersionNumber
        {
            get
            {
                try
                {
                    var baseKey = Registry.ClassesRoot;
                    var subKey = baseKey.OpenSubKey(@"Excel.Application\CurVer");
                    if (subKey != null)
                    {
                        switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                        {
                            // Excel 2003
                            case "EXCEL.APPLICATION.11":
                                return 11;

                            // Excel 2007
                            case "EXCEL.APPLICATION.12":
                                return 12;

                            // Excel 2010
                            case "EXCEL.APPLICATION.14":
                                return 14;

                            // Excel 2013
                            case "EXCEL.APPLICATION.15":
                                return 15;

                            default:
                                throw new OCExcelConfiguration("Could not determine Excel version");
                        }
                    }

                    throw new OCExcelConfiguration("Could not find registry key Excel.Application\\CurVer");
                }
                catch (Exception exception)
                {
                    throw new OCExcelConfiguration("Could not read registry to check Excel version", exception);
                }
            }
        }
        #endregion

        #region ExcelMaxRows
        /// <summary>
        /// Returns the maximum rows Excel supports
        /// </summary>
        /// <returns></returns>
        private static int ExcelMaxRows
        {
            get
            {
                const int excelMaxRowsFrom2003AndBelow = 65535;
                const int excelMaxRowsFrom2007AndUp = 1048576;

                switch (ExcelVersionNumber)
                {
                    // Excel 2007
                    case 12:
                    // Excel 2010
                    case 14:
                    // Excel 2013
                    case 15:
                        return excelMaxRowsFrom2007AndUp;

                    // Excel 2003 and older
                    default:
                        return excelMaxRowsFrom2003AndBelow;
                }
            }
        }
        #endregion

        #region ExcelColumnAddress
        /// <summary>
        /// Returns the column address for the given <paramref name="column"/>
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        private static string GetExcelColumnAddress(int column)
        {
            if (column <= 26)
                return System.Convert.ToChar(column + 64).ToString(CultureInfo.InvariantCulture);

            var div = column/26;
            var mod = column%26;
            if (mod != 0) return GetExcelColumnAddress(div) + GetExcelColumnAddress(mod);
            mod = 26;
            div--;

            return GetExcelColumnAddress(div) + GetExcelColumnAddress(mod);
        }
        #endregion

        #region GetExcelColumnNumber
        /// <summary>
        /// Returns the column number for the given <paramref name="columnAddress"/>
        /// </summary>
        /// <param name="columnAddress"></param>
        /// <returns></returns>
        // ReSharper disable once UnusedMember.Local
        private static int GetExcelColumnNumber(string columnAddress)
        {
            var digits = new int[columnAddress.Length];
            
            for (var i = 0; i < columnAddress.Length; ++i)
                digits[i] = System.Convert.ToInt32(columnAddress[i]) - 64;
            
            var mul = 1; 
            var res = 0;
            
            for (var pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }

            return res;
        }
        #endregion

        #region GetWorksheetPrintArea
        /// <summary>
        /// Placeholder for shape information
        /// </summary>
        private class ShapePosition
        {
            public int TopLeftColumn { get; private set; }
            public int TopLeftRow { get; private set; }
            public int BottomRightColumn { get; private set; }
            public int BottomRightRow { get; private set; }

            public ShapePosition(Excel.Shape shape)
            {
                var topLeftCell = shape.TopLeftCell;
                var bottomRightCell = shape.BottomRightCell;
                TopLeftRow = topLeftCell.Row;
                TopLeftColumn = topLeftCell.Column;
                BottomRightRow = bottomRightCell.Row;
                BottomRightColumn = bottomRightCell.Column;
                Marshal.ReleaseComObject(topLeftCell);
                Marshal.ReleaseComObject(bottomRightCell);
            }
        }

        /// <summary>
        /// Direction to search in merged cell
        /// </summary>
        private enum MergedCellSearchOrder
        {
            /// <summary>
            /// Search for first row in the merge area
            /// </summary>
            FirstRow,

            /// <summary>
            /// Search for first column in the merge area
            /// </summary>
            FirstColumn,

            /// <summary>
            /// Search for last row in the merge area
            /// </summary>
            LastRow,

            /// <summary>
            /// Search for last column in the merge area
            /// </summary>
            LastColumn
        }

        /// <summary>
        /// Checks if the given cell is merged and if so returns the last column or row from this merge.
        /// When the cell is not merged it just returns the cell
        /// </summary>
        /// <param name="range">The cell</param>
        /// <param name="searchOrder"><see cref="MergedCellSearchOrder"/></param>
        /// <returns></returns>
        private static int CheckForMergedCell(Excel.Range range, MergedCellSearchOrder searchOrder)
        {
            if (range == null)
                return 0;

            var result = 0;
            var mergeArea = range.MergeArea;

            switch (searchOrder)
            {
                case MergedCellSearchOrder.FirstRow:
                    result = range.Row;
                    if (mergeArea != null)
                        result = mergeArea.Row;
                    break;

                case MergedCellSearchOrder.FirstColumn:
                    result = range.Column;
                    if (mergeArea != null)
                        result = mergeArea.Column;
                    break;

                case MergedCellSearchOrder.LastRow:
                    result = range.Row;
                    if (mergeArea != null)
                        result += mergeArea.Rows.Count;
                    break;

                case MergedCellSearchOrder.LastColumn:
                    result = range.Column;
                    if (mergeArea != null)
                        result += mergeArea.Columns.Count;
                    break;
            }

            if (mergeArea != null)
                Marshal.ReleaseComObject(mergeArea);

            return result;
        }

        /// <summary>
        /// Figures out the used cell range. This are the cell's that contain any form of text and 
        /// returns this range. An empty range will be returned when there are shapes used on a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private static string GetWorksheetPrintArea(Excel._Worksheet worksheet)
        {
            var firstColumn = 1;
            var firstRow = 1;

            var shapesPosition = new List<ShapePosition>();

            // We can't use this method when there are shapes on a sheet so
            // we return an empty string
            var shapes = worksheet.Shapes;
            if (shapes.Count > 0)
            {
                if (ExcelVersionNumber < 14)
                    return "shapes";

                // The shape TopLeftCell and BottomRightCell is only supported from Excel 2010 and up
                foreach (Excel.Shape shape in worksheet.Shapes)
                {
                    if (shape.AutoShapeType != MsoAutoShapeType.msoShapeMixed)
                        shapesPosition.Add(new ShapePosition(shape));

                    Marshal.ReleaseComObject(shape);
                }

                Marshal.ReleaseComObject(shapes);
            }

            var range = worksheet.Cells[1, 1] as Excel.Range;
            if (range == null || range.Value == null)
            {
                if (range != null)
                    Marshal.ReleaseComObject(range);

                var firstCellByColumn = worksheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByColumns);
                var foundByFirstColumn = false;
                if (firstCellByColumn != null)
                {
                    foundByFirstColumn = true;
                    firstColumn = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstColumn);
                    firstRow = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstRow);
                    Marshal.ReleaseComObject(firstCellByColumn);
                }

                // Search the first used cell row wise
                var firstCellByRow = worksheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows);
                if (firstCellByRow == null)
                    return string.Empty;

                if (foundByFirstColumn)
                {
                    if (firstCellByRow.Column < firstColumn) firstColumn = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstColumn);
                    if (firstCellByRow.Row < firstRow) firstRow = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstRow);
                }
                else
                {
                    firstColumn = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstColumn);
                    firstRow = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstRow);
                }

                Marshal.ReleaseComObject(firstCellByRow);
            }

            foreach (var shapePosition in shapesPosition)
            {
                if (shapePosition.TopLeftColumn < firstColumn)
                    firstColumn = shapePosition.TopLeftColumn;

                if (shapePosition.TopLeftRow < firstRow)
                    firstRow = shapePosition.TopLeftRow;
            }

            var lastColumn = firstColumn;
            var lastRow = firstRow;

            var lastCellByColumn =
                worksheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByColumns,
                    SearchDirection: Excel.XlSearchDirection.xlPrevious);

            if (lastCellByColumn != null)
            {
                lastColumn = lastCellByColumn.Column;
                lastRow = lastCellByColumn.Row;
                Marshal.ReleaseComObject(lastCellByColumn);
            }

            var lastCellByRow =
                worksheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlPrevious);

            if (lastCellByRow != null)
            {
                if (lastCellByRow.Column > lastColumn) lastColumn = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastColumn);
                if (lastCellByRow.Row > lastRow) lastRow = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastRow);

                var protection = worksheet.Protection;
                if (!worksheet.ProtectContents || protection.AllowDeletingRows)
                {
                    var previousLastCellByRow =
                        worksheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows,
                            SearchDirection: Excel.XlSearchDirection.xlPrevious,
                            After: lastCellByRow);

                    Marshal.ReleaseComObject(lastCellByRow);

                    if (previousLastCellByRow != null)
                    {
                        var previousRow = CheckForMergedCell(previousLastCellByRow, MergedCellSearchOrder.LastRow);
                        Marshal.ReleaseComObject(previousLastCellByRow);

                        if (previousRow < lastRow - 2)
                        {
                            var rangeToDelete =
                                worksheet.Range[GetExcelColumnAddress(firstColumn) + (previousRow + 1) + ":" +
                                                GetExcelColumnAddress(lastColumn) + (lastRow - 2)];

                            rangeToDelete.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            Marshal.ReleaseComObject(rangeToDelete);
                            lastRow = previousRow + 2;
                        }
                    }

                    Marshal.ReleaseComObject(protection);
                }
            }

            foreach (var shapePosition in shapesPosition)
            {
                if (shapePosition.BottomRightColumn > lastColumn)
                    lastColumn = shapePosition.BottomRightColumn;

                if (shapePosition.BottomRightRow > lastRow)
                    lastRow = shapePosition.BottomRightRow;
            }

            return GetExcelColumnAddress(firstColumn) + firstRow + ":" +
                   GetExcelColumnAddress(lastColumn) + lastRow;
        }
        #endregion

        #region SetWorkSheetPaperSize
        private class ExcelPaperSize
        {
            public Excel.XlPaperSize PaperSize { get; private set; }
            public Excel.XlPageOrientation Orientation { get; private set; }

            public ExcelPaperSize(Excel.XlPaperSize paperSize, Excel.XlPageOrientation orientation)
            {
                PaperSize = paperSize;
                Orientation = orientation;
            }
        }

        /// <summary>
        /// Returns the total number of vertical pagebreaks in the print area
        /// </summary>
        /// <param name="pageBreaks"></param>
        /// <returns></returns>
        private static int CountVerticalPageBreaks(Excel.VPageBreaks pageBreaks)
        {
            var result = 0;

            try
            {
                foreach (Excel.VPageBreak pageBreak in pageBreaks)
                {
                    if (pageBreak.Extent == Excel.XlPageBreakExtent.xlPageBreakPartial)
                        result += 1;

                    Marshal.ReleaseComObject(pageBreak);
                }
            }
            catch (COMException)
            {
                result = pageBreaks.Count;
            }

            return result;
        }

        /// <summary>
        /// This method wil figure out the optimal paper size to use and sets it
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="printArea"></param>
        private static void SetWorkSheetPaperSize(Excel._Worksheet worksheet, string printArea)
        {
            var pageSetup = worksheet.PageSetup;

            try
            {
                pageSetup.Order = Excel.XlOrder.xlOverThenDown;

                var paperSizes = new List<ExcelPaperSize>
            {
                new ExcelPaperSize(Excel.XlPaperSize.xlPaperA4, Excel.XlPageOrientation.xlPortrait),
                new ExcelPaperSize(Excel.XlPaperSize.xlPaperA4, Excel.XlPageOrientation.xlLandscape),
                new ExcelPaperSize(Excel.XlPaperSize.xlPaperA3, Excel.XlPageOrientation.xlLandscape),
                new ExcelPaperSize(Excel.XlPaperSize.xlPaperA3, Excel.XlPageOrientation.xlPortrait)
            };

                var zoomRatios = new List<int> { 100, 95, 90, 85, 80, 75 };
                pageSetup.PrintArea = printArea;
                pageSetup.LeftHeader = worksheet.Name;
                
                foreach (var paperSize in paperSizes)
                {
                    var exitfor = false;
                    pageSetup.PaperSize = paperSize.PaperSize;
                    pageSetup.Orientation = paperSize.Orientation;
                    worksheet.ResetAllPageBreaks();

                    foreach (var zoomRatio in zoomRatios)
                    {
                        pageSetup.Zoom = false;
#pragma warning disable 219
                        // Yes these page counts look lame, but so is Excel 2010 in not updating
                        // the pages collection otherwis. We need to call the count methods to
                        // make this code work
                        var pages = pageSetup.Pages.Count;
                        pageSetup.Zoom = zoomRatio;
                        // ReSharper disable once RedundantAssignment
                        pages = pageSetup.Pages.Count;
#pragma warning restore 219

                        if (CountVerticalPageBreaks(worksheet.VPageBreaks) == 0)
                        {
                            exitfor = true;
                            break;
                        }
                    }

                    if (exitfor)
                        break;
                }
            }
            finally
            {
                if (pageSetup != null)
                    Marshal.ReleaseComObject(pageSetup);
            }

        }
        #endregion

        #region ConvertWithExcel
        /// <summary>
        /// Converts a Excel sheet to PDF
        /// </summary>
        /// <param name="inputFile">The Excel input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private static void ConvertWithExcel(string inputFile, string outputFile)
        {
            // We only need to perform this check if we are running on a server
            if (NativeMethods.IsWindowsServer())
                CheckIfSystemProfileDesktopDirectoryExists();

            CheckIfPrinterIsInstalled();

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
                    AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                    PrintCommunication = true,
                    Visible = true
                };

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

                // We cannot determine a print area when the document is marked as final so we remove this
                workbook.Final = false;

                var usedSheets = 0;

                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    try
                    {
                        var protection = sheet.Protection;
                        if (!sheet.ProtectContents || protection.AllowFormattingColumns)
                            sheet.Columns.AutoFit();

                        Marshal.ReleaseComObject(protection);
                    }
                    catch (COMException)
                    {
                        // Do nothing, this sometimes failes and there is nothing we can do about it
                    }

                    var printArea = GetWorksheetPrintArea(sheet);

                    switch (printArea)
                    {
                        case "shapes":
                            SetWorkSheetPaperSize(sheet, string.Empty);
                            usedSheets += 1;
                            break;

                        case "":
                            break;

                        default:
                            SetWorkSheetPaperSize(sheet, printArea);
                            usedSheets += 1;
                            break;
                    }

                    Marshal.ReleaseComObject(sheet);
                }

                // It is not possible in Excel to export an empty workbook
                if (usedSheets != 0)
                    workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile);
                else
                    throw new OCFileContainsNoData("The file '" + Path.GetFileName(inputFile) + "' contains no data"); 
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

            GC.Collect();
            GC.WaitForPendingFinalizers();
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
                        if (recordType != 0x2F) return false;
                        binaryReader.ReadUInt16();
                        var filePassRecord = new FilePassRecord(memoryStream);
                        var key = Biff8EncryptionKey.Create(filePassRecord.DocId);
                        return !key.Validate(filePassRecord.SaltData, filePassRecord.SaltHash);
                    }
                }
            }
            catch (OCExcelConfiguration)
            {
                // If we get an OCExcelConfiguration exception it means we have an unknown encryption
                // type so we return a false so that Excel itself can figure out if the file is password
                // protected
                return false;
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
                else if (line.Contains("'")) textQualifier = Excel.XlTextQualifier.xlTextQualifierSingleQuote;
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
                        var excelMaxRows = ExcelMaxRows;
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
                            Password: "dummypassword",
                            IgnoreReadOnlyRecommended: true,
                            AddToMru: false);

                }
            }
            catch (COMException comException)
            {
                if (comException.ErrorCode == -2146827284)
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");

                throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                                        "' could not be opened, error: " + ExceptionHelpers.GetInnerException(comException));
            }
            catch (Exception exception)
            {
                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' could not be opened, error: " +
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
