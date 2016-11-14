using System;
using System.IO;
using System.Runtime.InteropServices;
using ICSharpCode.SharpZipLib.Zip;
using OfficeConverter.Exceptions;
using OfficeConverter.Helpers;

/*
   Copyright 2014 - 2016 Kees van Spelde

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
                    if (Word.IsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    Word.Convert(inputFile, outputFile);
                    break;

                case ".ODT":
                    if (OpenDocumentFormatIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    Word.Convert(inputFile, outputFile);
                    break;

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
                    if (Excel.IsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");
                    Excel.Convert(inputFile, outputFile);
                    break;

                case ".CSV":
                    Excel.Convert(inputFile, outputFile);
                    break;

                case ".ODS":
                    if (OpenDocumentFormatIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

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
                    if (PowerPoint.IsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) + "' is password protected");

                    PowerPoint.Convert(inputFile, outputFile);
                    break;

                case ".ODP":
                    if (OpenDocumentFormatIsPasswordProtected(inputFile))
                        throw new OCFileIsPasswordProtected("The file '" + Path.GetFileName(inputFile) +
                                                            "' is password protected");

                    PowerPoint.Convert(inputFile, outputFile);
                    break;

                default:
                    throw new OCFileTypeNotSupported("The file '" + Path.GetFileName(inputFile) +
                                                     "' is not supported only, " + Environment.NewLine +
                                                     ".DOC, .DOT, .DOCM, .DOCX, .DOTM, .ODT, .RTF, .MHT, " + Environment.NewLine +
                                                     ".WPS, .WRI, .XLS, .XLT, .XLW, .XLSB, .XLSM, .XLSX, " + Environment.NewLine +
                                                     ".XLTM, .XLTX, .CSV, .ODS, .POT, .PPT, .PPS, .POTM, " + Environment.NewLine +
                                                     ".POTX, .PPSM, .PPSX, .PPTM, .PPTX, .ODP" + Environment.NewLine +
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
    }
}
