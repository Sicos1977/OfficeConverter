﻿using System;
using System.IO;
using System.Reflection;
using System.Xml;
using Microsoft.Extensions.Logging;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;
using PasswordProtectedChecker;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

//
// Converter.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2024 Magic-Sessions. (www.magic-sessions.com)
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
        ///     <see cref="Helpers.Logger"/>
        /// </summary>
        private Logger _logger;

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
        ///     <see cref="LibreOffice"/>
        /// </summary>
        private LibreOffice _libreOffice;

        /// <summary>
        ///     Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;
        #endregion

        #region Properties
        /// <summary>
        ///     When set then this directory is used to store temporary files
        /// </summary>
        public string TempDirectory { get; set; }

        /// <summary>
        ///     When set to <c>true</c> then the <see cref="TempDirectory"/>
        ///     will not be deleted when the extraction is done
        /// </summary>
        /// <remarks>
        ///     For debugging
        /// </remarks>
        public bool DoNotDeleteTempDirectory { get; set; }

        /// <summary>
        ///     When set then LibreOffice is used to do the conversion instead of Microsoft Office
        /// </summary>
        public bool UseLibreOffice { get; set; }

        /// <summary>
        /// Returns a reference to the LibreOffice class when it already exists or creates a new one
        /// when it doesn't
        /// </summary>
        private LibreOffice LibreOffice
        {
            get
            {
                if (_libreOffice != null)
                    return _libreOffice;

                _libreOffice = new LibreOffice(_logger);
                return _libreOffice;
            }
        }

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

                _word = new Word(_logger);
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
                {
                    _excel.TempDirectory = TempDirectory;
                    _excel.DoNotDeleteTempDirectory = DoNotDeleteTempDirectory;
                    return _excel;
                }

                _excel = new Excel(_logger);
                if (TempDirectory != null)
                {
                    _excel.TempDirectory = TempDirectory;
                    _excel.DoNotDeleteTempDirectory = DoNotDeleteTempDirectory;
                }

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

                _powerPoint = new PowerPoint(_logger);
                return _powerPoint;
            }
        }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets it's needed properties
        /// </summary>
        /// <param name="logger">When set then logging is written to this ILogger instance for all conversions at the Information log level. If
        ///     you want a separate log for each conversion then set the <see cref="ILogger"/> on the <see cref="Convert"/> method</param>
        /// <param name="instanceId">An unique id that can be used to identify the logging of the converter when
        ///     calling the code from multiple threads and writing all the logging to the same file</param>
        public Converter(ILogger logger = null, string instanceId = null)
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomainAssemblyResolve;
            _logger = new Logger(logger, instanceId);
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
        private void CheckFileNameAndOutputFolder(string inputFile, string outputFile)
        {
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentNullException(inputFile);

            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentNullException(outputFile);

            if (!File.Exists(inputFile))
            {
                var message = $"Could not find the input file '{inputFile}'";
                _logger?.WriteToLog(message);
                throw new FileNotFoundException(message);
            }

            var directoryInfo = new FileInfo(outputFile).Directory;
            if (directoryInfo == null) return;

            var outputFolder = directoryInfo.FullName;

            if (!Directory.Exists(outputFolder))
            {
                var message = $"The output folder '{outputFolder}' does not exist";
                _logger?.WriteToLog(message);
                throw new DirectoryNotFoundException(message);
            }
        }
        #endregion

        #region ThrowPasswordProtected
        private void ThrowPasswordProtected(string inputFile)
        {
            var message = $"The file '{Path.GetFileName(inputFile)}' is password protected";
            _logger?.WriteToLog(message);
            throw new OCFileIsPasswordProtected(message);
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts the <paramref name="inputFile" /> to PDF and saves it as the <paramref name="outputFile" />
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFile">The output file with full path</param>
        /// <param name="logger">>When set then logging is written to this ILogger instance at the Information log level</param>
        /// <param name="instanceId"></param>
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
        public void Convert(string inputFile, string outputFile, ILogger logger = null, string instanceId = null)
        {
            if (logger != null)
                _logger = new Logger(logger, instanceId);

            CheckFileNameAndOutputFolder(inputFile, outputFile);

            var extension = Path.GetExtension(inputFile);
            extension = extension?.ToUpperInvariant();

            switch (extension)
            {
                case ".DOC":
                case ".DOT":
                case ".DOCM":
                case ".DOCX":
                case ".DOTX":
                case ".DOTM":
                case ".ODT":
                {
                    var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                    if (result.Protected)
                        ThrowPasswordProtected(inputFile);

                    if (UseLibreOffice)
                        LibreOffice.Convert(inputFile,outputFile);
                    else
                        Word.Convert(inputFile, outputFile);

                    break;
                }

                case ".RTF":
                case ".MHT":
                case ".WPS":
                case ".WRI":
                    if (UseLibreOffice)
                        LibreOffice.Convert(inputFile, outputFile);
                    else
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
                case ".ODS":
                {
                    var result = _passwordProtectedChecker.IsFileProtected(inputFile);
                    if (result.Protected)
                        ThrowPasswordProtected(inputFile);

                    if (UseLibreOffice)
                        LibreOffice.Convert(inputFile,outputFile);
                    else
                        Excel.Convert(inputFile, outputFile);

                    break;
                }

                case ".CSV":
                    if (UseLibreOffice)
                        LibreOffice.Convert(inputFile,outputFile);
                    else
                        Excel.Convert(inputFile, outputFile);

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
                case ".ODP":
                {
                    if (UseLibreOffice)
                        LibreOffice.Convert(inputFile,outputFile);
                    else
                        PowerPoint.Convert(inputFile, outputFile);

                    break;
                }

                case ".XML":
                    var progId = GetProgId(inputFile);
                    if (!string.IsNullOrWhiteSpace(progId))
                    {
                        switch (progId)
                        {
                            case "Word":
                                Word.Convert(inputFile, outputFile);
                                break;

                            case "Excel":
                                Excel.Convert(inputFile, outputFile);
                                break;
                        }
                    }
                    else
                        goto default;

                    break;


                default:
                {
                    var message = $"The file '{Path.GetFileName(inputFile)}' " +
                                  $"is not supported only, {Environment.NewLine}" +
                                  $".DOC, .DOT, .DOCM, .DOCX, .DOTM, .XML (Word or Excel) .ODT, .RTF, .MHT, {Environment.NewLine}" +
                                  $".WPS, .WRI, .XLS, .XLT, .XLW, .XLSB, .XLSM, .XLSX, {Environment.NewLine}" +
                                  $".XLTM, .XLTX, .CSV, .ODS, .POT, .PPT, .PPS, .POTM, {Environment.NewLine}" +
                                   ".POTX, .PPSM, .PPSX, .PPTM, .PPTX and .ODP are supported";

                    _logger?.WriteToLog(message);
                    throw new OCFileTypeNotSupported(message);
                }
            }
        }
        #endregion

        #region GetProgId
        /// <summary>
        /// Returns the progId that is inside the XML or <c>null</c> when not found
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static string GetProgId(string fileName)
        {
            try
            {
                var doc = new XmlDocument();
                doc.Load(fileName);
                if (doc.HasChildNodes)
                {
                    var processingInstructions = doc.ChildNodes[1];
                    switch (processingInstructions.Value)
                    {
                        case "progid=\"Word.Document\"":
                            return "Word";
                        
                        case "progid=\"Excel.Sheet\"":
                            return "Excel";
                    }
                }
            }
            catch
            {
                return null;
            }

            return null;
        }
        #endregion

        #region CurrentDomainAssemblyResolve
        /// <summary>
        /// Event to resolve 32 or 64 bits dll
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly CurrentDomainAssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assemblyName = args.Name.Split(new[] {','}, 2)[0] + ".dll";
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CLI",
                Environment.Is64BitProcess ? "x64" : "x86", assemblyName);

            return File.Exists(path)
                ? Assembly.LoadFile(path)
                : null;
        }
        #endregion

        #region Dispose
        /// <summary>
        ///     Disposes all created office objects
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;

            AppDomain.CurrentDomain.AssemblyResolve -= CurrentDomainAssemblyResolve;

            if (_word != null)
            {
                _logger?.WriteToLog("Disposing Word object");
                _word.Dispose();
                _word = null;
            }

            if (_excel != null)
            {
                _logger?.WriteToLog("Disposing Excel object");
                _excel.Dispose();
                _excel = null;
            }

            if (_powerPoint != null)
            {
                _logger?.WriteToLog("Disposing PowerPoint object");
                _powerPoint.Dispose();
                _powerPoint = null;
            }

            if (_libreOffice != null)
            {
                _logger?.WriteToLog("Disposing LibreOffice object");
                _libreOffice.Dispose();
                _libreOffice = null;
            }

            _disposed = true;
        }
        #endregion
    }
}