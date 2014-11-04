using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CompoundFileStorage;
using CompoundFileStorage.Exceptions;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Biff8;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeConverter
{
    /// <summary>
    /// This class is used as a placeholder for al Excel related methods
    /// </summary>
    internal static class Excel
    {
        #region Private class ShapePosition
        /// <summary>
        /// Placeholder for shape information
        /// </summary>
        private class ShapePosition
        {
            public int TopLeftColumn { get; private set; }
            public int TopLeftRow { get; private set; }
            public int BottomRightColumn { get; private set; }
            public int BottomRightRow { get; private set; }

            public ShapePosition(ExcelInterop.Shape shape)
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
        #endregion

        #region Private class ExcelPaperSize
        /// <summary>
        /// Placeholder for papersize and orientation information
        /// </summary>
        private class ExcelPaperSize
        {
            public ExcelInterop.XlPaperSize PaperSize { get; private set; }
            public ExcelInterop.XlPageOrientation Orientation { get; private set; }

            public ExcelPaperSize(ExcelInterop.XlPaperSize paperSize, ExcelInterop.XlPageOrientation orientation)
            {
                PaperSize = paperSize;
                Orientation = orientation;
            }
        }
        #endregion

        #region Private enum MergedCellSearchOrder
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
        #endregion

        #region Fields
        private static readonly int VersionNumber;
        private static readonly int MaxRows;
        #endregion

        #region Constructor
        /// <summary>
        /// This constructor is called the first time when the <see cref="Convert"/> or
        /// <see cref="FileIsPasswordProtected"/> method is called. Some checks are done to
        /// see if all requirements for a succesfull conversion are there.
        /// </summary>
        static Excel()
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
                            VersionNumber = 11;
                            break;

                            // Excel 2007
                        case "EXCEL.APPLICATION.12":
                            VersionNumber = 12;
                            break;

                            // Excel 2010
                        case "EXCEL.APPLICATION.14":
                            VersionNumber = 14;
                            break;

                            // Excel 2013
                        case "EXCEL.APPLICATION.15":
                            VersionNumber = 15;
                            break;

                        default:
                            throw new OCExcelConfiguration("Could not determine Excel version");
                    }
                }
                else
                    throw new OCExcelConfiguration("Could not find registry key ExcelInterop.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCExcelConfiguration("Could not read registry to check Excel version", exception);
            }

            const int excelMaxRowsFrom2003AndBelow = 65535;
            const int excelMaxRowsFrom2007AndUp = 1048576;

            switch (VersionNumber)
            {
                // Excel 2007
                case 12:
                // Excel 2010
                case 14:
                // Excel 2013
                case 15:
                    MaxRows = excelMaxRowsFrom2007AndUp;
                    break;

                // Excel 2003 and older
                default:
                    MaxRows = excelMaxRowsFrom2003AndBelow;
                    break;
            }

            CheckIfSystemProfileDesktopDirectoryExists();
            CheckIfPrinterIsInstalled();
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

        #region GetColumnAddress
        /// <summary>
        /// Returns the column address for the given <paramref name="column"/>
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        private static string GetColumnAddress(int column)
        {
            if (column <= 26)
                return System.Convert.ToChar(column + 64).ToString(CultureInfo.InvariantCulture);

            var div = column / 26;
            var mod = column % 26;
            if (mod != 0) return GetColumnAddress(div) + GetColumnAddress(mod);
            mod = 26;
            div--;

            return GetColumnAddress(div) + GetColumnAddress(mod);
        }
        #endregion

        #region GetColumnNumber
        /// <summary>
        /// Returns the column number for the given <paramref name="columnAddress"/>
        /// </summary>
        /// <param name="columnAddress"></param>
        /// <returns></returns>
        // ReSharper disable once UnusedMember.Local
        private static int GetColumnNumber(string columnAddress)
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
        /// Checks if the given cell is merged and if so returns the last column or row from this merge.
        /// When the cell is not merged it just returns the cell
        /// </summary>
        /// <param name="range">The cell</param>
        /// <param name="searchOrder"><see cref="MergedCellSearchOrder"/></param>
        /// <returns></returns>
        private static int CheckForMergedCell(ExcelInterop.Range range, MergedCellSearchOrder searchOrder)
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
        private static string GetWorksheetPrintArea(ExcelInterop._Worksheet worksheet)
        {
            var firstColumn = 1;
            var firstRow = 1;

            var shapesPosition = new List<ShapePosition>();

            // We can't use this method when there are shapes on a sheet so
            // we return an empty string
            var shapes = worksheet.Shapes;
            if (shapes.Count > 0)
            {
                if (VersionNumber < 14)
                    return "shapes";

                // The shape TopLeftCell and BottomRightCell is only supported from Excel 2010 and up
                foreach (ExcelInterop.Shape shape in worksheet.Shapes)
                {
                    if (shape.AutoShapeType != MsoAutoShapeType.msoShapeMixed)
                        shapesPosition.Add(new ShapePosition(shape));

                    Marshal.ReleaseComObject(shape);
                }

                Marshal.ReleaseComObject(shapes);
            }

            var range = worksheet.Cells[1, 1] as ExcelInterop.Range;
            if (range == null || range.Value == null)
            {
                if (range != null)
                    Marshal.ReleaseComObject(range);

                var firstCellByColumn = worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByColumns);
                var foundByFirstColumn = false;
                if (firstCellByColumn != null)
                {
                    foundByFirstColumn = true;
                    firstColumn = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstColumn);
                    firstRow = CheckForMergedCell(firstCellByColumn, MergedCellSearchOrder.FirstRow);
                    Marshal.ReleaseComObject(firstCellByColumn);
                }

                // Search the first used cell row wise
                var firstCellByRow = worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows);
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
                worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByColumns,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious);

            if (lastCellByColumn != null)
            {
                lastColumn = lastCellByColumn.Column;
                lastRow = lastCellByColumn.Row;
                Marshal.ReleaseComObject(lastCellByColumn);
            }

            var lastCellByRow =
                worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious);

            if (lastCellByRow != null)
            {
                if (lastCellByRow.Column > lastColumn) lastColumn = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastColumn);
                if (lastCellByRow.Row > lastRow) lastRow = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastRow);

                var protection = worksheet.Protection;
                if (!worksheet.ProtectContents || protection.AllowDeletingRows)
                {
                    var previousLastCellByRow =
                        worksheet.Cells.Find("*", SearchOrder: ExcelInterop.XlSearchOrder.xlByRows,
                            SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious,
                            After: lastCellByRow);

                    Marshal.ReleaseComObject(lastCellByRow);

                    if (previousLastCellByRow != null)
                    {
                        var previousRow = CheckForMergedCell(previousLastCellByRow, MergedCellSearchOrder.LastRow);
                        Marshal.ReleaseComObject(previousLastCellByRow);

                        if (previousRow < lastRow - 2)
                        {
                            var rangeToDelete =
                                worksheet.Range[GetColumnAddress(firstColumn) + (previousRow + 1) + ":" +
                                                GetColumnAddress(lastColumn) + (lastRow - 2)];

                            rangeToDelete.Delete(ExcelInterop.XlDeleteShiftDirection.xlShiftUp);
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

            return GetColumnAddress(firstColumn) + firstRow + ":" +
                   GetColumnAddress(lastColumn) + lastRow;
        }
        #endregion

        #region CountVerticalPageBreaks
        /// <summary>
        /// Returns the total number of vertical pagebreaks in the print area
        /// </summary>
        /// <param name="pageBreaks"></param>
        /// <returns></returns>
        private static int CountVerticalPageBreaks(ExcelInterop.VPageBreaks pageBreaks)
        {
            var result = 0;

            try
            {
                foreach (ExcelInterop.VPageBreak pageBreak in pageBreaks)
                {
                    if (pageBreak.Extent == ExcelInterop.XlPageBreakExtent.xlPageBreakPartial)
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
        #endregion
        
        #region SetWorkSheetPaperSize
        /// <summary>
        /// This method wil figure out the optimal paper size to use and sets it
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="printArea"></param>
        private static void SetWorkSheetPaperSize(ExcelInterop._Worksheet worksheet, string printArea)
        {
            var pageSetup = worksheet.PageSetup;

            try
            {
                pageSetup.Order = ExcelInterop.XlOrder.xlOverThenDown;

                var paperSizes = new List<ExcelPaperSize>
            {
                new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlPortrait),
                new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlLandscape),
                new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlLandscape),
                new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlPortrait)
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

        #region Convert
        /// <summary>
        /// Converts a Excel sheet to PDF
        /// </summary>
        /// <param name="inputFile">The Excel input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        internal static void Convert(string inputFile, string outputFile)
        {
            // We only need to perform this check if we are running on a server
            if (NativeMethods.IsWindowsServer())
                CheckIfSystemProfileDesktopDirectoryExists();

            CheckIfPrinterIsInstalled();

            ExcelInterop.Application excel = null;
            ExcelInterop.Workbook workbook = null;
            string tempFileName = null;

            try
            {
                excel = new ExcelInterop.ApplicationClass
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

                workbook = Open(excel, inputFile, extension, false);

                // We cannot determine a print area when the document is marked as final so we remove this
                workbook.Final = false;

                var usedSheets = 0;

                foreach (ExcelInterop.Worksheet sheet in workbook.Sheets)
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
                    workbook.ExportAsFixedFormat(ExcelInterop.XlFixedFormatType.xlTypePDF, outputFile);
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

        #region FileIsPasswordProtected
        /// <summary>
        /// Returns true when the Excel file is password protected
        /// </summary>
        /// <param name="inputFile">The Excel file to check</param>
        /// <returns></returns>
        /// <exception cref="OCFileIsCorrupt">Raised when the file is corrupt</exception>
        internal static bool FileIsPasswordProtected(string inputFile)
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

        #region GetCsvSeperator
        /// <summary>
        /// Returns the seperator and textqualifier that is used in the CSV file
        /// </summary>
        /// <param name="inputFile">The inputfile</param>
        /// <param name="separator">The separator that is used</param>
        /// <param name="textQualifier">The text qualifier</param>
        /// <returns></returns>
        private static void GetCsvSeperator(string inputFile, out string separator, out ExcelInterop.XlTextQualifier textQualifier)
        {
            separator = string.Empty;
            textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierNone;

            using (var streamReader = new StreamReader(inputFile))
            {
                var line = string.Empty;
                while (string.IsNullOrEmpty(line))
                    line = streamReader.ReadLine();

                if (line.Contains(";")) separator = ";";
                else if (line.Contains(",")) separator = ",";
                else if (line.Contains("\t")) separator = "\t";
                else if (line.Contains(" ")) separator = " ";

                if (line.Contains("\"")) textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierDoubleQuote;
                else if (line.Contains("'")) textQualifier = ExcelInterop.XlTextQualifier.xlTextQualifierSingleQuote;
            }
        }
        #endregion

        #region Open
        /// <summary>
        /// Opens the <paramref name="inputFile"/> and returns it as an <see cref="ExcelInterop.Workbook"/> object
        /// </summary>
        /// <param name="excel">The <see cref="ExcelInterop.Application"/></param>
        /// <param name="inputFile">The file to open</param>
        /// <param name="extension">The file extension</param>
        /// <param name="repairMode">When true the <paramref name="inputFile"/> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile"/> has to many rows</exception>
        private static ExcelInterop.Workbook Open(ExcelInterop._Application excel,
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
                        var excelMaxRows = MaxRows;
                        if (count > excelMaxRows)
                            throw new OCCsvFileLimitExceeded("The input CSV file has more then " + excelMaxRows +
                                                             " rows, the installed Excel version supports only " +
                                                             excelMaxRows + " rows");

                        string separator;
                        ExcelInterop.XlTextQualifier textQualifier;

                        GetCsvSeperator(inputFile, out separator, out textQualifier);

                        switch (separator)
                        {
                            case ";":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited,
                                    textQualifier, true, false, true);
                                return excel.ActiveWorkbook;

                            case ",":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, true);
                                return excel.ActiveWorkbook;

                            case "\t":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, true);
                                return excel.ActiveWorkbook;

                            case " ":
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, false, true);
                                return excel.ActiveWorkbook;

                            default:
                                excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, true);
                                return excel.ActiveWorkbook;
                        }

                    default:

                        if (repairMode)
                            return excel.Workbooks.Open(inputFile, false, true,
                                Password: "dummypassword",
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false,
                                CorruptLoad: ExcelInterop.XlCorruptLoad.xlRepairFile);

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

                return Open(excel, inputFile, extension, true);
            }
        }
        #endregion
    }
}
