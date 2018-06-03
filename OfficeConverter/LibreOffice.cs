using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;

// Libreoffice assemblies
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.frame;

namespace OfficeConverter
{
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
        #region Start
        /// <summary>
        /// Checks if LibreOffice is started and if not starts it
        /// </summary>
        private static void Start()
        {
            var getProcess = Process.GetProcessesByName("soffice.exe");
            if (getProcess.Length != 0)
                throw new InvalidProgramException("OpenOffice not found.  Is OpenOffice installed?");

            if (getProcess.Length > 0)
                return;

            var startProcess = new Process
            {
                StartInfo =
                {
                    Arguments = "-headless -nofirststartwizard",
                    FileName = "soffice.exe",
                    CreateNoWindow = true
                }
            };

            if (!startProcess.Start())
                throw new InvalidProgramException("OpenOffice failed to start.");
        }
        #endregion

        public static void ConvertToPdf(string inputFile, string outputFile)
        {
            if (GetFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + inputFile);

            Start();

            var bootstrap = Bootstrap.bootstrap();
            // ReSharper disable once SuspiciousTypeConversion.Global
            var remoteFactory = (XMultiServiceFactory)bootstrap.getServiceManager();
            var aLoader = (XComponentLoader) remoteFactory.createInstance("com.sun.star.frame.Desktop");

            XComponent xComponent = null;

            try
            {
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
