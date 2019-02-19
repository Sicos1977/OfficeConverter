using System;
using System.IO;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using PasswordProtectedChecker;

//
// Converter.cs
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
    ///     With this class an Microsoft Office document can be converted to PDF format. Microsoft Office 2007
    ///     (with PDF export plugin) or higher is needed.
    /// </summary>
    public class Converter : IDisposable
    {
        #region Fields
        /// <summary>
        ///     <see cref="Checker"/>
        /// </summary>
        private readonly Checker _passwordProtectedChecker = new Checker();

        /// <summary>
        ///     <see cref="Word"/>
        /// </summary>
        private Word _word;

        /// <summary>
        ///     <see cref="Excel"/>
        /// </summary>
        private Excel _excel;

        /// <summary>
        ///     <see cref="PowerPoint"/>
        /// </summary>
        private PowerPoint _powerPoint;

        /// <summary>
        ///     Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;
        #endregion

        #region Properties
        /// <summary>
        ///     An unique id that can be used to identify the logging of the converter when
        ///     calling the code from multiple threads and writing all the logging to the same file
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public string InstanceId
        {
            get => Logger.InstanceId;
            set => Logger.InstanceId = value;
        }

        /// <summary>
        ///     When set then this directory is used to store temporary files
        /// </summary>
        public string TempDirectory { get; set; }

        /// <summary>
        /// Returns a reference to the Word class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private Word Word
        {
            get
            {
                if (_word != null)
                    return _word;

                _word = new Word();
                return _word;
            }
        }

        /// <summary>
        /// Returns a reference to the Excel class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private Excel Excel
        {
            get
            {
                if (_excel != null)
                    return _excel;

                _excel = new Excel();
                if (TempDirectory != null)
                    _excel.TempDirectory = TempDirectory;

                return _excel;
            }
        }


        /// <summary>
        /// Returns a reference to the PowerPoint class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private PowerPoint PowerPoint
        {
            get
            {
                if (_powerPoint != null)
                    return _powerPoint;

                _powerPoint = new PowerPoint();
                return _powerPoint;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets it's needed properties
        /// </summary>
        /// <param name="logStream">When set then logging is written to this stream for all conversions. If
        /// you want a separate log for each conversion then set the logstream on the <see cref="Convert"/> method</param>
        public Converter(Stream logStream = null)
        {
            Logger.LogStream = logStream;
        }
        #endregion

        #region CheckFileNameAndOutputFolder
        /// <summary>
        ///     Checks if the <paramref name="inputFile" /> and the folder where the <paramref name="outputFile" /> is written
        ///     exists
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        /// <exception cref="ArgumentNullException">
        ///     Raised when the <paramref name="inputFile" /> or <paramref name="outputFile" />
        ///     is null or empty
        /// </exception>
        /// <exception cref="FileNotFoundException">Raised when the <paramref name="inputFile" /> does not exists</exception>
        /// <exception cref="DirectoryNotFoundException">
        ///     Raised when the folder where the <paramref name="outputFile" /> is written
        ///     does not exists
        /// </exception>
        private static void CheckFileNameAndOutputFolder(string inputFile, string outputFile)
        {
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentNullException(inputFile);

            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentNullException(outputFile);

            if (!File.Exists(inputFile))
            {
                var message = $"Could not find the input file '{inputFile}'";
                Logger.WriteToLog(message);
                throw new FileNotFoundException(message);
            }

            var directoryInfo = new FileInfo(outputFile).Directory;
            if (directoryInfo == null) return;

            var outputFolder = directoryInfo.FullName;

            if (!Directory.Exists(outputFolder))
            {
                var message = $"The output folder '{outputFolder}' does not exist";
                Logger.WriteToLog(message);
                throw new DirectoryNotFoundException(message);
            }
        }
        #endregion

        #region ThrowPasswordProtected
        private void ThrowPasswordProtected(string inputFile)
        {
            var message = "The file '" + Path.GetFileName(inputFile) +
                          "' is password protected";
            Logger.WriteToLog(message);
            throw new OCFileIsPasswordProtected(message);
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts the <paramref name="inputFile" /> to PDF and saves it as the <paramref name="outputFile" />
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFile">The output file with full path</param>
        /// <param name="logStream">When set then logging is written to this stream</param>
        /// <exception cref="ArgumentNullException">
        ///     Raised when the <paramref name="inputFile" /> or <paramref name="outputFile" />
        ///     is null or empty
        /// </exception>
        /// <exception cref="FileNotFoundException">Raised when the <paramref name="inputFile" /> does not exist</exception>
        /// <exception cref="DirectoryNotFoundException">
        ///     Raised when the folder where the <paramref name="outputFile" /> is written
        ///     does not exists
        /// </exception>
        /// <exception cref="OCFileIsCorrupt">Raised when the <paramref name="inputFile" /> is corrupt</exception>
        /// <exception cref="OCFileTypeNotSupported">Raised when the <paramref name="inputFile" /> is not supported</exception>
        /// <exception cref="OCFileIsPasswordProtected">Raised when the <paramref name="inputFile" /> is password protected</exception>
        /// <exception cref="OCCsvFileLimitExceeded">Raised when a CSV <paramref name="inputFile" /> has to many rows</exception>
        /// <exception cref="OCFileContainsNoData">Raised when the Microsoft Office file contains no actual data</exception>
        public void Convert(string inputFile, string outputFile, Stream logStream = null)
        {
            if (logStream != null)
                Logger.LogStream = logStream;

            CheckFileNameAndOutputFolder(inputFile, outputFile);

            var extension = Path.GetExtension(inputFile);
            extension = extension?.ToUpperInvariant();

            switch (extension)
            {
                case ".DOC":
                case ".DOT":
                case ".DOCM":
                case ".DOCX":
                case ".DOTM":
                case ".ODT":
                {
                    var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                    if (result.Protected)
                        ThrowPasswordProtected(inputFile);

                    Word.Convert(inputFile, outputFile);
                    break;
                }

                case ".RTF":
                case ".MHT":
                case ".WPS":
                case ".WRI":
                    Word.Convert(inputFile, outputFile);
                    break;

                case ".XLS":
                case ".XLT":
                case ".XLW":
                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                {
                    var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                    if (result.Protected)
                        ThrowPasswordProtected(inputFile);

                    Excel.Convert(inputFile, outputFile);
                    break;
                }

                case ".CSV":
                    Excel.Convert(inputFile, outputFile);
                    break;

                case ".ODS":
                {
                    var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                    if (result.Protected)
                        ThrowPasswordProtected(inputFile);

                    Excel.Convert(inputFile, outputFile);
                    break;
                }

                case ".POT":
                case ".PPT":
                case ".PPS":
                case ".POTM":
                case ".POTX":
                case ".PPSM":
                case ".PPSX":
                case ".PPTM":
                case ".PPTX":
                {
                    var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                    if (result.Protected)
                        ThrowPasswordProtected(inputFile);

                    PowerPoint.Convert(inputFile, outputFile);
                    break;
                }

                case ".ODP":
                {
                    var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                    if (result.Protected)
                        ThrowPasswordProtected(inputFile);

                    PowerPoint.Convert(inputFile, outputFile);
                    break;
                }

                default:
                {
                    var message = "The file '" + Path.GetFileName(inputFile) +
                                  "' is not supported only, " + Environment.NewLine +
                                  ".DOC, .DOT, .DOCM, .DOCX, .DOTM, .ODT, .RTF, .MHT, " + Environment.NewLine +
                                  ".WPS, .WRI, .XLS, .XLT, .XLW, .XLSB, .XLSM, .XLSX, " + Environment.NewLine +
                                  ".XLTM, .XLTX, .CSV, .ODS, .POT, .PPT, .PPS, .POTM, " + Environment.NewLine +
                                  ".POTX, .PPSM, .PPSX, .PPTM, .PPTX, .ODP" + Environment.NewLine +
                                  " are supported";

                    Logger.WriteToLog(message);
                    throw new OCFileTypeNotSupported(message);
                }
            }
        }
        #endregion

        #region Dispose
        /// <summary>
        ///     Disposes all created office objects
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_word != null)
            {
                Logger.WriteToLog("Disposing Word object");
                _word.Dispose();
            }

            if (_excel != null)
            {
                Logger.WriteToLog("Disposing Excel object");
                _excel.Dispose();
            }

            if (_powerPoint != null)
            {
                Logger.WriteToLog("Disposing PowerPoint object");
                _powerPoint.Dispose();
            }
        }
        #endregion
    }
}