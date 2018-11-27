using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.util;

//
// LibreOffice.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2018 Magic-Sessions. (www.magic-sessions.com)
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
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace OfficeConverter
{
    /// <summary>
    /// This class is used as a placeholder for all Libre office related methods
    /// </summary>
    /// <remarks>
    /// - https://api.libreoffice.org/examples/examples.html
    /// - https://api.libreoffice.org/docs/install.html
    /// - https://www.libreoffice.org/download/download/
    /// </remarks>
    internal class LibreOffice
    {
        #region Fields
        private Process _libreOfficeProcess;
        private string _userFolder;
        private string _pipeName;
        #endregion

        #region Properties
        /// <summary>
        /// Returns the full path to LibreOffice, when not found <c>null</c> is returned
        /// </summary>
        private string GetInstallPath
        {
            get
            {
                using (var regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false))
                {
                    var installPath = (string) regkey?.GetValue(string.Empty);
                    return installPath;
                }
            }
        }
        #endregion

        #region Start
        /// <summary>
        /// Checks if LibreOffice is started and if not starts it
        /// </summary>
        private void Start()
        {
            var installPath = GetInstallPath;
            if (string.IsNullOrEmpty(installPath))
                throw new InvalidProgramException("LibreOffice not installed");

            var path = installPath.Replace('\\', '/');

            Environment.SetEnvironmentVariable("URE_BOOTSTRAP", "vnd.sun.star.pathname:" + path + "/fundamental.ini");
            var environmentPath = Environment.GetEnvironmentVariable("PATH");

            Environment.SetEnvironmentVariable("UNO_PATH", path, EnvironmentVariableTarget.Process);

            if (environmentPath != null && !environmentPath.Contains(path))
                Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + @";" + path,
                    EnvironmentVariableTarget.Process);

            var process = new Process
            {
                StartInfo =
                {
                    // -env:UserInstallation=file:///{_userFolder}
                    Arguments =
                        $"-invisible -nofirststartwizard -minimized -nologo -nolockcheck --accept=pipe,name={_pipeName};urp;StarOffice.ComponentContext",
                    FileName = installPath + @"\soffice.exe",
                    CreateNoWindow = true
                }
            };

            if (!process.Start())
                throw new InvalidProgramException("Could not start LibreOffice");

            _libreOfficeProcess = process;
        }
        #endregion

        #region ConvertToUrl
        /// <summary>
        /// Convert the give file path to the format LibreOffice needs
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public string ConvertToUrl(string file)
        {
            return $"file:///{file.Replace(@"\", "/")}";
        }
        #endregion

        #region ConvertToPdf
        /// <summary>
        /// Converts the given <paramref name="inputFile"/> to PDF format and saves it as <paramref name="outputFile"/>
        /// </summary>
        /// <param name="inputFile">The input file</param>
        /// <param name="outputFile">The output file</param>
        public void ConvertToPdf(string inputFile, string outputFile)
        {
            if (GetFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + inputFile);

            XComponent component = null;

            try
            {
                var guid = Guid.NewGuid().ToString().Replace("-", string.Empty);
                _userFolder = $"d:/{guid}";
                _pipeName = guid;
                //_pipeName = "keeshispipe";
                //Directory.CreateDirectory(_userFolder);

                Start();

                var localContext = Bootstrap.defaultBootstrap_InitialComponentContext();
                var localServiceManager = localContext.getServiceManager();
                var urlResolver = (XUnoUrlResolver) localServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext);
                XComponentContext remoteContext;

                var i = 0;

                while (true)
                {
                    try
                    {
                        remoteContext =
                            (XComponentContext) urlResolver.resolve(
                                $"uno:pipe,name={_pipeName};urp;StarOffice.ComponentContext");

                        break;
                    }
                    catch (System.Exception exception)
                    {
                        if (i == 20 || !exception.Message.Contains("couldn't connect to pipe")) throw;
                        Thread.Sleep(100);
                        i++;
                    }
                }

                // ReSharper disable once SuspiciousTypeConversion.Global
                var remoteFactory = (XMultiServiceFactory) remoteContext.getServiceManager();
                var componentLoader = (XComponentLoader) remoteFactory.createInstance("com.sun.star.frame.Desktop"); 
                component = InitDocument(componentLoader, ConvertToUrl(inputFile), "_blank");

                // Save/export the document
                // http://herbertniemeyerblog.blogspot.com/2011/11/have-to-start-somewhere.html
                // https://forum.openoffice.org/en/forum/viewtopic.php?t=73098

                ExportToPdf(component, inputFile, outputFile);

                CloseDocument(component);
            }
            finally
            {
                component?.dispose();

                if (_libreOfficeProcess != null && !_libreOfficeProcess.HasExited)
                {
                    _libreOfficeProcess.Kill();

                    while(!_libreOfficeProcess.HasExited)
                        Thread.Sleep(100);

                    _libreOfficeProcess = null;
                }

                try
                {
                    if (!string.IsNullOrEmpty(_userFolder))
                        Directory.Delete(_userFolder, true);
                }
                catch
                {
                }
            }
        }
        #endregion

        #region InitDocument
        /// <summary>
        /// Creates a new document in LibreOffice and opens the given <paramref name="file"/>
        /// </summary>
        /// <param name="aLoader"></param>
        /// <param name="file"></param>
        /// <param name="target"></param>
        /// <returns></returns>
        private XComponent InitDocument(XComponentLoader aLoader, string file, string target)
        {
            var openProps = new PropertyValue[2];
            openProps[0] = new PropertyValue { Name = "Hidden", Value = new Any(true) };
            openProps[1] = new PropertyValue { Name = "ReadOnly", Value = new Any(true) };

            var xComponent = aLoader.loadComponentFromURL(
                file, target, 0,
                openProps);

            return xComponent;
        }
        #endregion

        #region SaveDocument
        /// <summary>
        /// Exports the loaded document to PDF format
        /// </summary>
        /// <param name="component"></param>
        /// <param name="sourceFile"></param>
        /// <param name="destinationFile"></param>
        private  void ExportToPdf(XComponent component, string sourceFile, string destinationFile)
        {
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
                Value = new Any(GetFilterType(sourceFile))
            };
            
            // Setting the flag for overwriting
            propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new Any(true) };

            var polymorphicType = PolymorphicType.GetType(
                typeof(PropertyValue[]),
                "unoidl.com.sun.star.beans.PropertyValue[]");

            propertyValues[2] = new PropertyValue { Name = "FilterData",  Value = new Any(polymorphicType, filterData) };
            
            // ReSharper disable once SuspiciousTypeConversion.Global
            ((XStorable)component).storeToURL(ConvertToUrl(destinationFile), propertyValues);
        }
        #endregion

        #region CloseDocument
        /// <summary>
        /// Closes the document and frees any used resources
        /// </summary>
        private void CloseDocument(XComponent component)
        {
            var closeable = (XCloseable)component;
            closeable?.close(false);
        }
        #endregion

        #region GetFilterType
        /// <summary>
        /// Returns the filter that is needed to convert the given <paramref name="fileName"/>,
        /// <c>null</c> is returned when the file cannot be converted
        /// </summary>
        /// <param name="fileName">The file to check</param>
        /// <returns></returns>
        public string GetFilterType(string fileName)
        {
            var extension = Path.GetExtension(fileName);

            switch (extension)
            {
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".odt":
                case ".wps":
                case ".wpd":
                    return "writer_pdf_Export";

                case ".xls":
                case ".xlsb":
                case ".xlsx":
                case ".xlsm":
                case ".ods":
                    return "calc_pdf_Export";

                case ".ppt":
                case ".pptx":
                case ".odp":
                    return "impress_pdf_Export";

                default:
                    return null;
            }
        }
        #endregion
    }
}
