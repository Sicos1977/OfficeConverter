using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
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
        private static void StartOpenOffice()
        {
            var ps = Process.GetProcessesByName("soffice.exe");
            if (ps.Length != 0)
                throw new InvalidProgramException("OpenOffice not found.  Is OpenOffice installed?");
            if (ps.Length > 0)
                return;
            var p = new Process
            {
                StartInfo =
                        {
                            Arguments = "-headless -nofirststartwizard",
                            FileName = "soffice.exe",
                            CreateNoWindow = true
                        }
            };
            var result = p.Start();

            if (result == false)
                throw new InvalidProgramException("OpenOffice failed to start.");
        }

        public static void ConvertToPdf(string inputFile, string outputFile)
        {
            if (ConvertExtensionToFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + inputFile);

            StartOpenOffice();

            //Get a ComponentContext
            var xLocalContext = Bootstrap.bootstrap();
            //Get MultiServiceFactory
            var xRemoteFactory = (XMultiServiceFactory) xLocalContext.getServiceManager();
            //Get a CompontLoader
            var aLoader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
            //Load the sourcefile

            XComponent xComponent = null;
            try
            {
                xComponent = InitDocument(aLoader, PathConverter(inputFile), "_blank");
                //Wait for loading
                while (xComponent == null)
                    Thread.Sleep(1000);

                // save/export the document
                SaveDocument(xComponent, inputFile, PathConverter(outputFile));
            }
            finally
            {
                if (xComponent != null) xComponent.dispose();
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
            //// Setting the filter name
            propertyValues[0] = new PropertyValue
            {
                Name = "FilterName",
                Value = new Any(ConvertExtensionToFilterType(Path.GetExtension(sourceFile)))
            };
            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);
        }

        public static string ConvertExtensionToFilterType(string extension)
        {
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
    }
}
