using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Win32;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using ExcelInterop = Microsoft.Office.Interop.Excel;

//
// Excel.cs
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
    ///     This class is used as a placeholder for all Excel related methods
    /// </summary>
    internal class Excel : IDisposable
    {
        #region Private class ShapePosition
        /// <summary>
        ///     Placeholder for shape information
        /// </summary>
        private class ShapePosition
        {
            /// <summary>
            ///     Returns the top left column
            /// </summary>
            public int TopLeftColumn { get; }

            /// <summary>
            ///     Returns the top left row
            /// </summary>
            public int TopLeftRow { get; }

            /// <summary>
            ///     Returns the bottom right column
            /// </summary>
            public int BottomRightColumn { get; }

            /// <summary>
            ///     Returns the bottom right row
            /// </summary>
            public int BottomRightRow { get; }

            /// <summary>
            ///     Creates this object and sets it's needed properties
            /// </summary>
            /// <param name="shape">The shape object</param>
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
        ///     Placeholder for papersize and orientation information
        /// </summary>
        private class ExcelPaperSize
        {
            /// <summary>
            ///     Returns the papersize
            /// </summary>
            public ExcelInterop.XlPaperSize PaperSize { get; }

            /// <summary>
            ///     Returns the orientation
            /// </summary>
            public ExcelInterop.XlPageOrientation Orientation { get; }

            /// <summary>
            ///     Creates this object and sets it's needed properties
            /// </summary>
            /// <param name="paperSize">The papersize</param>
            /// <param name="orientation">The orientation</param>
            public ExcelPaperSize(ExcelInterop.XlPaperSize paperSize, ExcelInterop.XlPageOrientation orientation)
            {
                PaperSize = paperSize;
                Orientation = orientation;
            }
        }
        #endregion

        #region Private enum MergedCellSearchOrder
        /// <summary>
        ///     Direction to search in merged cells
        /// </summary>
        private enum MergedCellSearchOrder
        {
            /// <summary>
            ///     Search for first row in the merge area
            /// </summary>
            FirstRow,

            /// <summary>
            ///     Search for first column in the merge area
            /// </summary>
            FirstColumn,

            /// <summary>
            ///     Search for last row in the merge area
            /// </summary>
            LastRow,

            /// <summary>
            ///     Search for last column in the merge area
            /// </summary>
            LastColumn
        }
        #endregion

        #region Fields
        /// <summary>
        ///     Excel version number
        /// </summary>
        private readonly int _versionNumber;

        /// <summary>
        ///     Excel maximum rows
        /// </summary>
        private readonly int _maxRows;

        /// <summary>
        ///     Paper sizes to use when detecting optimal page size with the <see cref="SetWorkSheetPaperSize" /> method
        /// </summary>
        private readonly List<ExcelPaperSize> _paperSizes = new List<ExcelPaperSize>
        {
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlPortrait),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA4, ExcelInterop.XlPageOrientation.xlLandscape),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlLandscape),
            new ExcelPaperSize(ExcelInterop.XlPaperSize.xlPaperA3, ExcelInterop.XlPageOrientation.xlPortrait)
        };

        /// <summary>
        ///     Zoom ration to use when detecting optimal page size with the <see cref="SetWorkSheetPaperSize" /> method
        /// </summary>
        private readonly List<int> _zoomRatios = new List<int> {100, 95, 90, 85, 80, 75, 70};

        /// <summary>
        ///     <see cref="ExcelInterop.ApplicationClass"/>
        /// </summary>
        private ExcelInterop.ApplicationClass _excel;

        /// <summary>
        ///     A <see cref="Process" /> object to Excel
        /// </summary>
        private Process _excelProcess;

        /// <summary>
        ///     When set then this folder is used for temporary files
        /// </summary>
        private DirectoryInfo _tempDirectory;

        /// <summary>
        ///     Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;
        #endregion

        #region Properties
        /// <summary>
        ///     When set then this directory is used to store temporary files
        /// </summary>
        /// <exception cref="DirectoryNotFoundException">Raised when the given directory does not exists</exception>
        public string TempDirectory
        {
            get => _tempDirectory.FullName;
            set
            {
                if (!Directory.Exists(value))
                    throw new DirectoryNotFoundException($"The directory '{value}' does not exists");

                _tempDirectory = new DirectoryInfo(Path.Combine(value, Guid.NewGuid().ToString()));
            }
        }

        /// <summary>
        ///     Returns a reference to the temp directory
        /// </summary>
        private DirectoryInfo GetTempDirectory
        {
            get
            {
                if (_tempDirectory == null)
                    _tempDirectory = new DirectoryInfo(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));

                if (!_tempDirectory.Exists)
                    _tempDirectory.Create();

                return _tempDirectory;
            }
        }

        /// <summary>
        ///     Returns <c>true</c> when Excel is running
        /// </summary>
        /// <returns></returns>
        private bool IsExcelRunning
        {
            get
            {
                if (_excelProcess == null)
                    return false;

                _excelProcess.Refresh();
                return !_excelProcess.HasExited;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     This constructor checks to see if all requirements for a successful conversion are here.
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the registry could not be read to determine Excel version</exception>
        internal Excel()
        {
            Logger.WriteToLog("Checking what version of Excel is installed");

            try
            {
                var baseKey = Registry.ClassesRoot;
                var subKey = baseKey.OpenSubKey(@"Excel.Application\CurVer");
                if (subKey != null)
                    switch (subKey.GetValue(string.Empty).ToString().ToUpperInvariant())
                    {
                        // Excel 2003
                        case "EXCEL.APPLICATION.11":
                            _versionNumber = 11;
                            Logger.WriteToLog("Excel 2003 is installed");
                            break;

                        // Excel 2007
                        case "EXCEL.APPLICATION.12":
                            _versionNumber = 12;
                            Logger.WriteToLog("Excel 2007 is installed");
                            break;

                        // Excel 2010
                        case "EXCEL.APPLICATION.14":
                            _versionNumber = 14;
                            Logger.WriteToLog("Excel 2010 is installed");
                            break;

                        // Excel 2013
                        case "EXCEL.APPLICATION.15":
                            _versionNumber = 15;
                            Logger.WriteToLog("Excel 2013 is installed");
                            break;

                        // Excel 2016
                        case "EXCEL.APPLICATION.16":
                            _versionNumber = 16;
                            Logger.WriteToLog("Excel 2016 is installed");
                            break;

                        // Excel 2019
                        case "EXCEL.APPLICATION.17":
                            _versionNumber = 17;
                            Logger.WriteToLog("Excel 2019 is installed");
                            break;

                        default:
                            throw new OCConfiguration("Could not determine Excel version");
                    }
                else
                    throw new OCConfiguration("Could not find registry key Excel.Application\\CurVer");
            }
            catch (Exception exception)
            {
                throw new OCConfiguration("Could not read registry to check Excel version", exception);
            }

            const int excelMaxRowsFrom2003AndBelow = 65535;
            const int excelMaxRowsFrom2007AndUp = 1048576;

            switch (_versionNumber)
            {
                // Excel 2007
                case 12:
                // Excel 2010
                case 14:
                // Excel 2013
                case 15:
                // Excel 2016
                case 16:
                // Excel 2019
                case 17:
                    _maxRows = excelMaxRowsFrom2007AndUp;
                    break;

                // Excel 2003 and older
                default:
                    _maxRows = excelMaxRowsFrom2003AndBelow;
                    break;
            }

            Logger.WriteToLog($"Setting maximum Excel rows to {_maxRows}");

            // We only need to perform this check if we are running on a server
            if (NativeMethods.IsWindowsServer())
                CheckIfSystemProfileDesktopDirectoryExists();
            
            CheckIfPrinterIsInstalled();
        }
        #endregion

        #region StartExcel
        /// <summary>
        ///     Starts Excel
        /// </summary>
        private void StartExcel()
        {
            if (IsExcelRunning)
            {
                Logger.WriteToLog($"Excel is already running on PID {_excelProcess.Id}... skipped");
                return;
            }

            Logger.WriteToLog("Starting Excel");

            _excel = new ExcelInterop.ApplicationClass
            {
                ScreenUpdating = false,
                DisplayAlerts = false,
                DisplayDocumentInformationPanel = false,
                DisplayRecentFiles = false,
                DisplayScrollBars = false,
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                PrintCommunication = true, // DO NOT REMOVE THIS LINE, NO NEVER EVER ... DON'T EVEN TRY IT
                Visible = false
            };

            ProcessHelpers.GetWindowThreadProcessId(_excel.Hwnd, out var processId);
            _excelProcess = Process.GetProcessById(processId);

            Logger.WriteToLog($"Excel started with process id {_excelProcess.Id}");
        }
        #endregion

        #region StopExcel
        /// <summary>
        ///     Stops Excel
        /// </summary>
        private void StopExcel()
        {
            if (IsExcelRunning)
            {
                Logger.WriteToLog("Stopping Excel");
                _excel.Quit();

                var counter = 0;

                // Give Excel 2 seconds to close
                while (counter < 2000)
                {
                    if (!IsExcelRunning) break;
                    counter++;
                    Thread.Sleep(1);
                }

                if (IsExcelRunning)
                {
                    Logger.WriteToLog($"Excel did not shutdown gracefully... killing it on process id {_excelProcess.Id}");
                    _excelProcess.Kill();
                    _excelProcess = null;
                    Logger.WriteToLog("Excel process killed");
                }
                else
                    Logger.WriteToLog("Excel stopped");
            }

            if (_excel != null)
            {
                Marshal.ReleaseComObject(_excel);
                _excel = null;
            }

            _excelProcess = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region CheckIfSystemProfileDesktopDirectoryExists
        /// <summary>
        ///     If you want to run this code on a server then the following folders must exist, if they don't
        ///     then you can't use Excel to convert files to PDF
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when the needed directory could not be created</exception>
        private void CheckIfSystemProfileDesktopDirectoryExists()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                var x64DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"SysWOW64\config\systemprofile\desktop");

                Logger.WriteToLog($"Checking if system profile desktop directory exists in '{x64DesktopPath}'");

                if (!Directory.Exists(x64DesktopPath))
                    try
                    {
                        Directory.CreateDirectory(x64DesktopPath);
                        Logger.WriteToLog("Directory did not exist ... created it");
                    }
                    catch (Exception exception)
                    {
                        throw new OCConfiguration("Can't create folder '" + x64DesktopPath +
                                                  "' Excel needs this folder to work on a server, error: " +
                                                  ExceptionHelpers.GetInnerException(exception));
                    }
            }
            else
            {
                var x86DesktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"System32\config\systemprofile\desktop");

                Logger.WriteToLog($"Checking if system profile desktop directory exists in '{x86DesktopPath}'");

                if (!Directory.Exists(x86DesktopPath))
                    try
                    {
                        Directory.CreateDirectory(x86DesktopPath);
                        Logger.WriteToLog("Directory did not exist ... created it");
                    }
                    catch (Exception exception)
                    {
                        throw new OCConfiguration("Can't create folder '" + x86DesktopPath +
                                                  "' Excel needs this folder to work on a server, error: " +
                                                  ExceptionHelpers.GetInnerException(exception));
                    }
            }
        }
        #endregion

        #region CheckIfPrinterIsInstalled
        /// <summary>
        ///     Excel needs a default printer to export to PDF, this method will check if there is one
        /// </summary>
        /// <exception cref="OCConfiguration">Raised when an default printer does not exists</exception>
        private void CheckIfPrinterIsInstalled()
        {
            Logger.WriteToLog("Excel needs a printer to convert sheets to pdf ... checking if a printer exists");

            var result = false;

            PrinterSettings.StringCollection installedPrinters;

            try
            {
                installedPrinters = PrinterSettings.InstalledPrinters;
            }
            catch (Win32Exception win32Exception)
            {
                throw new OCConfiguration($"Printer spooler service not enabled, error: {ExceptionHelpers.GetInnerException(win32Exception)}");
            }

            foreach (string printerName in installedPrinters)
            {
                // Retrieve the printer settings.
                var printer = new PrinterSettings {PrinterName = printerName};

                // Check that this is a valid printer.
                // (This step might be required if you read the printer name
                // from a user-supplied value or a registry or configuration file
                // setting.)
                if (printer.IsValid)
                {
                    Logger.WriteToLog($"A valid printer '{printer.PrinterName}' is found");
                    result = true;
                    break;
                }
            }

            if (!result)
                throw new OCConfiguration("There is no default printer installed, Excel needs one to export to PDF");
        }
        #endregion

        #region GetColumnAddress
        /// <summary>
        ///     Returns the column address for the given <paramref name="column" />
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        private string GetColumnAddress(int column)
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
        ///     Returns the column number for the given <paramref name="columnAddress" />
        /// </summary>
        /// <param name="columnAddress"></param>
        /// <returns></returns>
        // ReSharper disable once UnusedMember.Local
        private int GetColumnNumber(string columnAddress)
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

        #region CheckForMergedCell
        /// <summary>
        ///     Checks if the given cell is merged and if so returns the last column or row from this merge.
        ///     When the cell is not merged it just returns the cell
        /// </summary>
        /// <param name="range">The cell</param>
        /// <param name="searchOrder">
        ///     <see cref="MergedCellSearchOrder" />
        /// </param>
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
                    result = mergeArea.Row;
                    break;

                case MergedCellSearchOrder.FirstColumn:
                    result = mergeArea.Column;
                    break;

                case MergedCellSearchOrder.LastRow:
                {
                    result = range.Row;
                    var entireRow = range.EntireRow;

                    for (var i = 1; i < range.Column; i++)
                    {
                        var cell = (ExcelInterop.Range) entireRow.Cells[i];
                        var cellMergeArea = cell.MergeArea;
                        var cellMergeAreaRows = cellMergeArea.Rows;
                        var count = cellMergeAreaRows.Count;

                        Marshal.ReleaseComObject(cellMergeAreaRows);
                        Marshal.ReleaseComObject(cellMergeArea);
                        Marshal.ReleaseComObject(cell);

                        var tempResult = result;

                        if (count > 1 && range.Row + count > tempResult)
                            tempResult = result + count;

                        result = tempResult;
                    }

                    Marshal.ReleaseComObject(entireRow);

                    break;
                }

                case MergedCellSearchOrder.LastColumn:
                {
                    result = range.Column;
                    var columns = mergeArea.Columns;
                    var count = columns.Count;

                    if (count > 1)
                        result += count;

                    Marshal.ReleaseComObject(columns);

                    break;
                }
            }

            if (mergeArea != null)
                Marshal.ReleaseComObject(mergeArea);

            return result;
        }
        #endregion

        #region GetWorksheetPrintArea
        /// <summary>
        ///     Figures out the used cell range. This are the cell's that contain any form of text and
        ///     returns this range. An empty range will be returned when there are shapes used on a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private string GetWorksheetPrintArea(ExcelInterop._Worksheet worksheet)
        {
            var firstColumn = 1;
            var firstRow = 1;

            var shapesPosition = new List<ShapePosition>();

            // We can't use this method when there are shapes on a sheet so
            // we return an empty string
            var shapes = worksheet.Shapes;
            if (shapes.Count > 0)
            {
                if (_versionNumber < 14)
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
            if (range?.Value == null)
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
                    if (firstCellByRow.Column < firstColumn)
                        firstColumn = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstColumn);
                    if (firstCellByRow.Row < firstRow)
                        firstRow = CheckForMergedCell(firstCellByRow, MergedCellSearchOrder.FirstRow);
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
                if (lastCellByRow.Column > lastColumn)
                    lastColumn = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastColumn);

                if (lastCellByRow.Row > lastRow)
                    lastRow = CheckForMergedCell(lastCellByRow, MergedCellSearchOrder.LastRow);

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
        ///     Returns the total number of vertical pagebreaks in the print area
        /// </summary>
        /// <param name="pageBreaks"></param>
        /// <returns></returns>
        private int CountVerticalPageBreaks(ExcelInterop.VPageBreaks pageBreaks)
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
        ///     This method wil figure out the optimal paper size to use and sets it
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="printArea"></param>
        private void SetWorkSheetPaperSize(ExcelInterop._Worksheet worksheet, string printArea)
        {
            Logger.WriteToLog($"Detecting optimal paper size for sheet {worksheet.Name} with print area '{printArea}'");

            var pageSetup = worksheet.PageSetup;
            var pages = pageSetup.Pages;

            pageSetup.PrintArea = printArea;
            pageSetup.LeftHeader = worksheet.Name;

            var pageCount = pages.Count;

            if (pageCount == 1)
                return;

            try
            {
                pageSetup.Order = ExcelInterop.XlOrder.xlOverThenDown;

                foreach (var paperSize in _paperSizes)
                {
                    var exitfor = false;
                    pageSetup.PaperSize = paperSize.PaperSize;
                    pageSetup.Orientation = paperSize.Orientation;
                    worksheet.ResetAllPageBreaks();

                    foreach (var zoomRatio in _zoomRatios)
                    {
                        // Yes these page counts look lame, but so is Excel 2010 in not updating
                        // the pages collection otherwise. We need to call the count methods to
                        // make this code work
                        pageSetup.Zoom = zoomRatio;
                        // ReSharper disable once RedundantAssignment
                        pageCount = pages.Count;

                        if (CountVerticalPageBreaks(worksheet.VPageBreaks) == 0)
                        {
                            exitfor = true;
                            break;
                        }
                    }

                    if (exitfor)
                        break;
                }

                Logger.WriteToLog($"Paper size set to '{pageSetup.PaperSize}', orientation to '{pageSetup.Orientation}' and zoom ratio to '{pageSetup.Zoom}'");
            }
            finally
            {
                Marshal.ReleaseComObject(pages);
                Marshal.ReleaseComObject(pageSetup);
            }
        }
        #endregion

        #region SetChartPaperSize
        /// <summary>
        ///     This method wil set the papersize for a chart
        /// </summary>
        /// <param name="chart"></param>
        private void SetChartPaperSize(ExcelInterop._Chart chart)
        {
            Logger.WriteToLog($"Setting paper site for chart '{chart.Name}' to A4 landscape");

            var pageSetup = chart.PageSetup;
            var pages = pageSetup.Pages;

            try
            {
                pageSetup.LeftHeader = chart.Name;
                pageSetup.PaperSize = ExcelInterop.XlPaperSize.xlPaperA4;
                pageSetup.Orientation = ExcelInterop.XlPageOrientation.xlLandscape;
            }
            finally
            {
                Marshal.ReleaseComObject(pages);
                Marshal.ReleaseComObject(pageSetup);
            }
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts an Excel sheet to PDF
        /// </summary>
        /// <param name="inputFile">The Excel input file</param>
        /// <param name="outputFile">The PDF output file</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile" /> has to many rows</exception>
        internal void Convert(string inputFile, string outputFile)
        {
            DeleteResiliencyKeys();

            ExcelInterop.Workbook workbook = null;

            try
            {
                StartExcel();

                var extension = Path.GetExtension(inputFile);
                if (string.IsNullOrWhiteSpace(extension))
                    extension = string.Empty;

                if (extension.ToUpperInvariant() == ".CSV")
                {
                    var tempFileName = Path.Combine(GetTempDirectory.FullName, Guid.NewGuid() + ".txt");

                    // Yes this look somewhat weird but we have to change the extension if we want to handle
                    // CSV files with different kind of separators. Otherwhise Excel will always overrule whatever
                    // setting we make to open a file
                    Logger.WriteToLog($"Copying CSV file '{inputFile}' to temporary file '{tempFileName}' and setting that one as the input file");
                    File.Copy(inputFile, tempFileName);
                    inputFile = tempFileName;
                }

                workbook = OpenWorkbook(inputFile, extension, false);

                // We cannot determine a print area when the document is marked as final so we remove this
                workbook.Final = false;

                // Fix for "This command is not available in a shared workbook."
                if (workbook.MultiUserEditing)
                {
                    var tempFileName = Path.Combine(GetTempDirectory.FullName, Guid.NewGuid() + Path.GetExtension(inputFile));
                    Logger.WriteToLog($"Excel file '{inputFile}' is in 'multi user editing' mode saving it to temporary file '{tempFileName}' to set it to exclusive mode");
                    workbook.SaveAs(tempFileName, AccessMode: ExcelInterop.XlSaveAsAccessMode.xlExclusive);
                }

                var usedSheets = 0;

                var activeWindow = _excel.ActiveWindow;

                if (activeWindow == null)
                {
                    const string message = "There is no window active in Excel";
                    Logger.WriteToLog(message);
                    throw new OCFileContainsNoData(message);
                }

                foreach (var sheetObject in workbook.Sheets)
                {
                    switch (sheetObject)
                    {
                        // Invisible sheets will not be converted... they are not visible
                        case ExcelInterop.Worksheet sheet when sheet.Visible != ExcelInterop.XlSheetVisibility.xlSheetVisible:
                            continue;

                        case ExcelInterop.Worksheet sheet:
                            var protection = sheet.Protection;

                            try
                            {
                                // ReSharper disable once RedundantCast
                                (sheet as ExcelInterop._Worksheet).Activate();
                                if (!sheet.ProtectContents || protection.AllowFormattingColumns)
                                {
                                    if (activeWindow.View != ExcelInterop.XlWindowView.xlPageLayoutView)
                                    {
                                        Logger.WriteToLog($"Auto fitting colums on sheet '{sheet.Name}'");
                                        sheet.Columns.AutoFit();
                                    }
                                }
                            }
                            catch (COMException)
                            {
                                // Do nothing, this sometimes failes and there is nothing we can do about it
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(protection);
                            }

                            var printArea = GetWorksheetPrintArea(sheet);
                            Logger.WriteToLog($"Print area for sheet {sheet.Name} set to '{printArea}'");

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
                            continue;
                    }

                    if (!(sheetObject is ExcelInterop.Chart chart)) continue;
                    SetChartPaperSize(chart);

                    Marshal.ReleaseComObject(activeWindow);
                    Marshal.ReleaseComObject(chart);
                }

                // It is not possible in Excel to export an empty workbook
                if (usedSheets != 0)
                {
                    Logger.WriteToLog($"Exporting worksheets to PDF file '{outputFile}'");
                    workbook.ExportAsFixedFormat(ExcelInterop.XlFixedFormatType.xlTypePDF, outputFile);
                    Logger.WriteToLog("Worksheets exported to PDF");
                }
                else
                {
                    const string message = "The file contains no data";
                    Logger.WriteToLog(message);
                    throw new OCFileContainsNoData(message);
                }
            }
            catch (Exception)
            {
                StopExcel();
                throw;
            }

            finally
            {
                CloseWorkbook(workbook);

                if (_tempDirectory != null)
                {
                    _tempDirectory.Refresh();
                    if (_tempDirectory.Exists)
                    {
                        Logger.WriteToLog($"Deleting temporary folder '{_tempDirectory.FullName}'");
                        _tempDirectory.Delete(true);
                    }
                }
            }
        }
        #endregion

        #region GetCsvSeperator
        /// <summary>
        ///     Returns the separator and text qualifier that is used in the CSV file
        /// </summary>
        /// <param name="inputFile">The input file</param>
        /// <param name="separator">The separator that is used</param>
        /// <param name="textQualifier">The text qualifier</param>
        /// <returns></returns>
        private static void GetCsvSeparator(string inputFile, out string separator,
            out ExcelInterop.XlTextQualifier textQualifier)
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

        #region OpenWorkbook
        /// <summary>
        ///     Opens the <paramref name="inputFile" /> and returns it as an <see cref="ExcelInterop.Workbook" /> object
        /// </summary>
        /// <param name="inputFile">The file to open</param>
        /// <param name="extension">The file extension</param>
        /// <param name="repairMode">When true the <paramref name="inputFile" /> is opened in repair mode</param>
        /// <returns></returns>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile" /> has to many rows</exception>
        private ExcelInterop.Workbook OpenWorkbook(string inputFile, string extension, bool repairMode)
        {
            Logger.WriteToLog($"Opening workbook '{inputFile}'{(repairMode ? " with repair mode" : string.Empty)}");

            try
            {
                switch (extension.ToUpperInvariant())
                {
                    case ".CSV":

                        var count = File.ReadLines(inputFile).Count();
                        var excelMaxRows = _maxRows;
                        if (count > excelMaxRows)
                            throw new OCCsvFileLimitExceeded("The input CSV file has more then " + excelMaxRows +
                                                             " rows, the installed Excel version supports only " +
                                                             excelMaxRows + " rows");

                        GetCsvSeparator(inputFile, out var separator, out var textQualifier);
                        Logger.WriteToLog($"Separator for CSV file set to '{separator}' and text qualifier to '{textQualifier}'");

                        switch (separator)
                        {
                            case ";":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited,
                                    textQualifier, true, false, true);
                                break;

                            case ",":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, true);
                                break;

                            case "\t":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, true);
                                break;

                            case " ":
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, false, false, true);
                                break;

                            default:
                                _excel.Workbooks.OpenText(inputFile, Type.Missing, Type.Missing,
                                    ExcelInterop.XlTextParsingType.xlDelimited, textQualifier,
                                    Type.Missing, false, true);
                                break;
                        }

                        Logger.WriteToLog("Workbook opened");
                        return _excel.ActiveWorkbook;

                    default:

                        ExcelInterop.Workbook workbook;

                        if (repairMode)
                        {
                            workbook = _excel.Workbooks.Open(inputFile, false, true,
                                Password: "dummy password",
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false,
                                CorruptLoad: ExcelInterop.XlCorruptLoad.xlRepairFile);

                        }
                        else
                        {
                            workbook = _excel.Workbooks.Open(inputFile, false, true,
                                Password: "dummy password",
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false);
                        }

                        Logger.WriteToLog("Workbook opened");
                        return workbook;
                }
            }
            catch (COMException comException)
            {
                if (comException.ErrorCode == -2146827284)
                    throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                        "' is password protected");

                throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                          "' could not be opened, error: " +
                                          ExceptionHelpers.GetInnerException(comException));
            }
            catch (Exception exception)
            {
                Logger.WriteToLog(
                    $"ERROR: Failed to open worksheet, exception: '{ExceptionHelpers.GetInnerException(exception)}'");

                if (repairMode)
                    throw new OCFileIsCorrupt("The file '" + Path.GetFileName(inputFile) +
                                              "' could not be opened, error: " +
                                              ExceptionHelpers.GetInnerException(exception));

                return OpenWorkbook(inputFile, extension, true);
            }
        }
        #endregion

        #region CloseWorkbook
        /// <summary>
        ///     Closes the opened workbook and releases any allocated resources
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        private void CloseWorkbook(ExcelInterop.Workbook workbook)
        {
            if (workbook == null) return;
            Logger.WriteToLog("Closing workbook");
            workbook.Saved = true;
            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            Logger.WriteToLog("Workbook closed");
        }
        #endregion

        #region DeleteResiliencyKeys
        /// <summary>
        ///     This method will delete the automatic created Resiliency key. Excel uses this registry key
        ///     to make entries to corrupted workbooks. If there are to many entries under this key Excel will
        ///     get slower and slower to start. To prevent this we just delete this key when it exists
        /// </summary>
        private void DeleteResiliencyKeys()
        {
            Logger.WriteToLog("Deleting Excel resiliency keys from the registry");

            try
            {
                // HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Excel\Resiliency\DocumentRecovery
                var key = $@"Software\Microsoft\Office\{_versionNumber}.0\Excel\Resiliency";

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
        ///     Disposes the running <see cref="_excel" />
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            StopExcel();
        }
        #endregion
    }
}