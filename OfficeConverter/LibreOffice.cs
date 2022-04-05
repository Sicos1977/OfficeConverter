using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Microsoft.Win32;
using OfficeConverter.Helpers;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.util;
using Exception = System.Exception;

//
// LibreOffice.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2022 Magic-Sessions. (www.magic-sessions.com)
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
    ///     This class is used as a placeholder for all Libre office related methods
    /// </summary>
    /// <remarks>
    ///     - https://api.libreoffice.org/examples/examples.html
    ///     - https://api.libreoffice.org/docs/install.html
    ///     - https://www.libreoffice.org/download/download/
    /// </remarks>
    internal class LibreOffice : IDisposable
    {
        #region Fields
        /// <summary>
        ///     A <see cref="Process" /> object to LibreOffice
        /// </summary>
        private Process _libreOfficeProcess;

        /// <summary>
        ///     <see cref="XComponentLoader"/>
        /// </summary>
        private XComponentLoader _componentLoader;
        
        /// <summary>
        ///     <see cref="Logger"/>
        /// </summary>
        private readonly Logger _logger;

        /// <summary>
        ///     Keeps track is we already disposed our resources
        /// </summary>
        private bool _disposed;
        #endregion

        #region Properties
        /// <summary>
        ///     Returns the full path to LibreOffice, when not found <c>null</c> is returned
        /// </summary>
        private string GetInstallPath
        {
            get
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var regKey64 = hklm.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false))
                {
                    var installPath = (string)regKey64?.GetValue(string.Empty);

                    if (installPath != null)
                        return installPath;

                    using (var regKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false))
                    {
                        installPath = (string)regKey?.GetValue(string.Empty);
                        return installPath;
                    }
                }
            }
        }

        #region Properties
        /// <summary>
        ///     Returns <c>true</c> when LibreOffice is running
        /// </summary>
        /// <returns></returns>
        private bool IsLibreOfficeRunning
        {
            get
            {
                if (_libreOfficeProcess == null)
                    return false;

                _libreOfficeProcess.Refresh();
                return !_libreOfficeProcess.HasExited;
            }
        }
        #endregion
        #endregion

        #region Constructor
        internal LibreOffice(Logger logger)
        {
            _logger = logger;
        }
        #endregion

        #region StartLibreOffice
        /// <summary>
        ///     Checks if LibreOffice is started and if not starts it
        /// </summary>
        private void StartLibreOffice()
        {
            if (IsLibreOfficeRunning)
            {
                _logger?.WriteToLog($"LibreOffice is already running on PID {_libreOfficeProcess.Id}... skipped");
                return;
            }

            var installPath = GetInstallPath;
            if (string.IsNullOrEmpty(installPath))
                throw new InvalidProgramException("LibreOffice is not installed");

            var path = installPath.Replace('\\', '/');

            var ureBootStrap = $"vnd.sun.star.pathname:{path}/fundamental.ini";
            _logger?.WriteToLog($"Setting environment variable URE_BOOTSTRAP to '{ureBootStrap}'");
            Environment.SetEnvironmentVariable("URE_BOOTSTRAP", $"vnd.sun.star.pathname:{path}/fundamental.ini", EnvironmentVariableTarget.Process);
            
            var environmentPath = Environment.GetEnvironmentVariable("PATH");
            _logger?.WriteToLog($"Setting environment variable UNO_PATH to '{path}'");
            Environment.SetEnvironmentVariable("UNO_PATH", path, EnvironmentVariableTarget.Process);

            if (environmentPath != null && !environmentPath.Contains(path))
            {
                _logger?.WriteToLog($"Adding '{path}' to PATH environment variable");
                Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + @";" + path,
                    EnvironmentVariableTarget.Process);
            }

            _logger?.WriteToLog("Starting LibreOffice");
            
            var pipeName = Guid.NewGuid().ToString().Replace("-", string.Empty);

            var process = new Process
            {
                StartInfo =
                {
                    // -env:UserInstallation=file:///{_userFolder}
                    Arguments = $"-invisible -nofirststartwizard -minimized -nologo -nolockcheck --accept=pipe,name={pipeName};urp;StarOffice.ComponentContext",
                    FileName = installPath + @"\soffice.exe",
                    CreateNoWindow = true
                }
            };

            if (!process.Start())
                throw new InvalidProgramException("Could not start LibreOffice");

            _libreOfficeProcess = process;

            _logger?.WriteToLog($"LibreOffice started with process id {process.Id}");

            OpenLibreOfficePipe(pipeName);
        }
        #endregion

        #region OpenLibreOfficePipe
        /// <summary>
        ///     Opens a pipe to LibreOffice
        /// </summary>
        /// <param name="pipeName"></param>
        private void OpenLibreOfficePipe(string pipeName)
        {
            var localContext = Bootstrap.defaultBootstrap_InitialComponentContext();
            var localServiceManager = localContext.getServiceManager();
            var urlResolver = (XUnoUrlResolver)localServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext);
            XComponentContext remoteContext;

            var i = 0;

            _logger?.WriteToLog($"Connecting to LibreOffice with pipe '{pipeName}'");

            while (true)
                try
                {
                    remoteContext = (XComponentContext)urlResolver.resolve($"uno:pipe,name={pipeName};urp;StarOffice.ComponentContext");
                    _logger?.WriteToLog("Connected to LibreOffice");
                    break;
                }
                catch (Exception exception)
                {
                    if (i == 20 || !exception.Message.Contains("couldn't connect to pipe")) 
                        throw;

                    Thread.Sleep(100);
                    i++;
                }

            // ReSharper disable once SuspiciousTypeConversion.Global
            var remoteFactory = (XMultiServiceFactory)remoteContext.getServiceManager();
            _componentLoader = (XComponentLoader)remoteFactory.createInstance("com.sun.star.frame.Desktop");
        }
        #endregion

        #region StopLibreOffice
        /// <summary>
        ///     Stops LibreOffice
        /// </summary>
        private void StopLibreOffice()
        {
            if (IsLibreOfficeRunning)
            {
                _logger?.WriteToLog($"LibreOffice did not close gracefully, closing it by killing it's process on id {_libreOfficeProcess.Id}");
                _libreOfficeProcess.Kill();
                _libreOfficeProcess = null;
                _logger?.WriteToLog("LibreOffice process killed");
            }
            else
                _logger?.WriteToLog($"LibreOffice {(_libreOfficeProcess != null ? $"with process id {_libreOfficeProcess.Id} " : string.Empty)}already exited");

            _libreOfficeProcess = null;
        }
        #endregion

        #region ConvertToUrl
        /// <summary>
        ///     Convert the give file path to the format LibreOffice needs
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private string ConvertToUrl(string file)
        {
            return $"file:///{file.Replace(@"\", "/")}";
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts the given <paramref name="inputFile" /> to PDF format and saves it as <paramref name="outputFile" />
        /// </summary>
        /// <param name="inputFile">The input file</param>
        /// <param name="outputFile">The output file</param>
        public void Convert(string inputFile, string outputFile)
        {
            if (GetFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException($"Unknown file type '{Path.GetFileName(inputFile)}' for LibreOffice");

            StartLibreOffice();

            var component = InitDocument(_componentLoader, ConvertToUrl(inputFile), "_blank");

            // Save/export the document
            // http://herbertniemeyerblog.blogspot.com/2011/11/have-to-start-somewhere.html
            // https://forum.openoffice.org/en/forum/viewtopic.php?t=73098

            ExportToPdf(component, inputFile, outputFile);

            CloseDocument(component);
        }
        #endregion

        #region InitDocument
        /// <summary>
        ///     Creates a new document in LibreOffice and opens the given <paramref name="inputFile" />
        /// </summary>
        /// <param name="aLoader"></param>
        /// <param name="inputFile"></param>
        /// <param name="target"></param>
        /// <returns></returns>
        private XComponent InitDocument(XComponentLoader aLoader, string inputFile, string target)
        {
            _logger?.WriteToLog($"Loading document '{inputFile}'");

            var openProps = new PropertyValue[2];
            openProps[0] = new PropertyValue { Name = "Hidden", Value = new Any(true) };
            openProps[1] = new PropertyValue { Name = "ReadOnly", Value = new Any(true) };

            var xComponent = aLoader.loadComponentFromURL(
                inputFile, target, 0,
                openProps);

            _logger?.WriteToLog("Document loaded");

            return xComponent;
        }
        #endregion

        #region ExportToPdf
        /// <summary>
        ///     Exports the loaded document to PDF format
        /// </summary>
        /// <param name="component"></param>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        private void ExportToPdf(XComponent component, string inputFile, string outputFile)
        {
            _logger?.WriteToLog($"Exporting document to PDF file '{outputFile}'");

            var propertyValues = new PropertyValue[3];
            var filterData = new PropertyValue[5];

            filterData[0] = new PropertyValue
            {
                Name = "UseLosslessCompression",
                Value = new Any(false)
            };

            filterData[1] = new PropertyValue
            {
                Name = "Quality",
                Value = new Any(90)
            };

            filterData[2] = new PropertyValue
            {
                Name = "ReduceImageResolution",
                Value = new Any(true)
            };

            filterData[3] = new PropertyValue
            {
                Name = "MaxImageResolution",
                Value = new Any(300)
            };

            filterData[4] = new PropertyValue
            {
                Name = "ExportBookmarks",
                Value = new Any(false)
            };

            // Setting the filter name
            propertyValues[0] = new PropertyValue
            {
                Name = "FilterName",
                Value = new Any(GetFilterType(inputFile))
            };

            // Setting the flag for overwriting
            propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new Any(true) };

            var polymorphicType = PolymorphicType.GetType(typeof(PropertyValue[]), "unoidl.com.sun.star.beans.PropertyValue[]");

            propertyValues[2] = new PropertyValue { Name = "FilterData", Value = new Any(polymorphicType, filterData) };

            // ReSharper disable once SuspiciousTypeConversion.Global
            ((XStorable)component).storeToURL(ConvertToUrl(outputFile), propertyValues);

            _logger?.WriteToLog("Document exported to PDF");
        }
        #endregion

        #region CloseDocument
        /// <summary>
        ///     Closes the document and frees any used resources
        /// </summary>
        private void CloseDocument(XComponent component)
        {
            _logger?.WriteToLog("Closing document");
            var closeable = (XCloseable)component;
            closeable?.close(false);
            _logger?.WriteToLog("Document closed");
        }
        #endregion

        #region GetFilterType
        /// <summary>
        ///     Returns the filter that is needed to convert the given <paramref name="fileName" />,
        ///     <c>null</c> is returned when the file cannot be converted
        /// </summary>
        /// <param name="fileName">The file to check</param>
        /// <returns></returns>
        private string GetFilterType(string fileName)
        {
            var extension = Path.GetExtension(fileName);
            extension = extension?.ToUpperInvariant();

            switch (extension)
            {
                case ".DOC":
                case ".DOT":
                case ".DOCM":
                case ".DOCX":
                case ".DOTM":
                case ".ODT":
                case ".RTF":
                case ".MHT":
                case ".WPS":
                case ".WRI":
                    return "writer_pdf_Export";

                case ".XLS":
                case ".XLT":
                case ".XLW":
                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                    return "calc_pdf_Export";

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
                    return "impress_pdf_Export";

                default:
                    return null;
            }
        }
        #endregion

        #region Dispose
        /// <summary>
        ///     Disposes the running <see cref="_libreOfficeProcess" />
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            StopLibreOffice();
            _disposed = true;
        }
        #endregion
    }
}