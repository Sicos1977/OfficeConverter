using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.bridge;

// Libreoffice assemblies
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.uno;

namespace OfficeConverter
{
    /*


        XComponentContext xLocalContext = uno.util.Bootstrap.defaultBootstrap_InitialComponentContext();
        XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
        XUnoUrlResolver xUrlResolver = (XUnoUrlResolver) xLocalServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", xLocalContext);


        int i = 0;
        while (i < 20) {
            try
            {
                xContext = (XComponentContext)xUrlResolver.resolve(
                    "uno:pipe,name=officepipe1;urp;StarOffice.ComponentContext");
                if (xContext != null)
                    break;
            } catch (unoidl.com.sun.star.connection.NoConnectException) {
                System.Threading.Thread.Sleep(100);
            }
            i++;
        }
        if (xContext == null)
            return;

        XMultiServiceFactory xMsf = (XMultiServiceFactory) xContext.getServiceManager();

        Object desktop = xMsf.createInstance("com.sun.star.frame.Desktop");
        XComponentLoader xLoader = (XComponentLoader)desktop;
    */

    /// <summary>
    /// This class is used as a placeholder for all Libre office related methods
    /// </summary>
    /// <remarks>
    /// - https://api.libreoffice.org/examples/examples.html
    /// - https://api.libreoffice.org/docs/install.html
    /// - https://www.libreoffice.org/download/download/
    /// - https://github.com/dmazz55/Pdfvert/blob/master/Source/Pdfvert/Utilities/OpenOfficeUtility.cs
    /// </remarks>
    internal static class LibreOffice
    {
        #region Fields
        private static Process _libreOfficeProcess;
        #endregion

        #region Properties
        /// <summary>
        /// Returns the full path to LibreOffice, when not found <c>null</c> is returned
        /// </summary>
        private static string GetInstallPath
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
        private static void Start()
        {
            var installPath = GetInstallPath;
            if (string.IsNullOrEmpty(installPath))
                throw new InvalidProgramException("LibreOffice not installed");

            var path = installPath.Replace('\\', '/');

            Environment.SetEnvironmentVariable("URE_BOOTSTRAP", "vnd.sun.star.pathname:" + path + "/fundamental.ini");

            var process = new Process
            {
                StartInfo =
                {
                    Arguments =
                        "-nodefault -nologo -nofirststartwizard -accept=pipe,name=officepipe1;urp;StarOffice.ServiceManager",
                    FileName = installPath + @"\soffice.exe",
                    CreateNoWindow = true
                }
            };

            if (!process.Start())
                throw new InvalidProgramException("Could not start LibreOffice");

            _libreOfficeProcess = process;
        }
        #endregion

        public static void ConvertToPdf(string inputFile, string outputFile)
        {
            if (GetFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + inputFile);


            XComponent xComponent = null;

            try
            {
                Start();

                //var xLocalContext = Bootstrap.defaultBootstrap_InitialComponentContext();
                //var xLocalServiceManager = xLocalContext.getServiceManager();
                //var xUrlResolver = (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
                //    "com.sun.star.bridge.UnoUrlResolver", xLocalContext);

                var bootstrap = Bootstrap.defaultBootstrap_InitialComponentContext();
                // ReSharper disable once SuspiciousTypeConversion.Global
                var remoteFactory = (XMultiServiceFactory)bootstrap.getServiceManager();
                var aLoader = (XComponentLoader)remoteFactory.createInstance("com.sun.star.frame.Desktop");

                xComponent = InitDocument(aLoader, inputFile, "_blank");
                //Wait for loading
                //while (xComponent == null)
                //    Thread.Sleep(1000);

                // save/export the document
                SaveDocument(xComponent, inputFile, outputFile);
            }
            finally
            {
                xComponent?.dispose();

                if (_libreOfficeProcess != null)
                {
                    _libreOfficeProcess.Kill();
                    _libreOfficeProcess = null;
                }
            }
        }

        private static XComponent InitDocument(XComponentLoader aLoader, string file, string target)
        {
            var openProps = new PropertyValue[1];
            openProps[0] = new PropertyValue { Name = "Hidden", Value = new Any(true) };

            var xComponent = aLoader.loadComponentFromURL(
                file, target, 0,
                openProps);

            return xComponent;
        }

        private static void SaveDocument(XComponent xComponent, string sourceFile, string destinationFile)
        {
            var propertyValues = new PropertyValue[2];
            // Setting the flag for overwriting
            propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new Any(true) };
            // Setting the filter name
            propertyValues[0] = new PropertyValue
            {
                Name = "FilterName",
                Value = new Any(GetFilterType(sourceFile))
            };
            // ReSharper disable once SuspiciousTypeConversion.Global
            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);
        }

        #region GetFilterType
        /// <summary>
        /// Returns the filter that is needed to convert the given <paramref name="fileName"/>,
        /// <c>null</c> is returned when the file cannot be converted
        /// </summary>
        /// <param name="fileName">The file to check</param>
        /// <returns></returns>
        public static string GetFilterType(string fileName)
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
