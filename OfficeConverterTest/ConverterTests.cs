using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeConverter;
using OfficeConverter.Exceptions;

/*
   Copyright 2014-2015 Kees van Spelde

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

namespace OfficeConverterTest
{
    [TestClass]
    public class ExtractionTests
    {
        private readonly List<string> _tempFolders = new List<string>();

        [TestMethod]
        [ExpectedException(typeof(OCFileTypeNotSupported))]
        public void FileTypeNotSupported()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\filetypenotsupported.txt", outputFile);
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsCorrupt))]
        public void FileIsCorrupt()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A corrupt compound document.doc", outputFile);
        }

        #region Microsoft Office Word tests
        [TestMethod]
        public void DocWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A DOC word document without embedded files.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocWith7EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A DOC word document with 7 embedded files.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void DocWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A DOC word document with password.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A DOCX word document without embedded files.docx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocxWith7EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A DOCX word document with 7 embedded files.docx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void DocxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A DOCX word document with password.docx", outputFile);
        }
        #endregion

        #region Microsoft Office Excel tests
        [TestMethod]
        public void XlsWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A XLS excel document without embedded files.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsWith2EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A XLS excel document with 2 embedded files.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void XlsWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A XLS excel document with password.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A XLSX excel document without embedded files.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsxWith2EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A XLSX excel document with 2 embedded files.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void XlsxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A XLSX excel document with password.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvSemicolonSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\Semicolon separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvCommaSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\Comma separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvSpaceSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\Space separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvTabSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\Tab separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }
        #endregion

        #region Microsoft Office PowerPoint tests
        [TestMethod]
        public void PptWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A PPT PowerPoint document without embedded files.ppt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void PptWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A PPT powerpoint document with 3 embedded files.ppt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void PptWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A PPT PowerPoint document with password.ppt", outputFile);
        }

        [TestMethod]
        public void PptxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A PPTX PowerPoint document without embedded files.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void PptxWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A PPTX powerpoint document with 3 embedded files.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void PptxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\A PPTX PowerPoint document with password.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }
        #endregion

        #region Open Office Writer tests
        [TestMethod]
        public void OdtWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\An ODT document without embedded files.odt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void OdtWith8EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\An ODT document with 8 embedded files.odt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void OdtWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\An ODT document with password.odt", outputFile);
        }
        #endregion
        
        #region Open Office Impress tests
        [TestMethod]
        public void OdpWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\An ODP document without embedded files.odp", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void OdpWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\An ODP document with 3 embedded files.odp", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void OdpWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert(GetCurrentDir() + "TestFiles\\An ODP document with password.odp", outputFile);
        }
        #endregion

        #region Helper methods
        /// <summary>
        /// Creates a new temporary folder and returns the path to it
        /// </summary>
        /// <returns></returns>
        private string CreateTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            _tempFolders.Add(tempDirectory);
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }

        private static void DeleteTemporaryFolder(string folder)
        {
            try
            {
                if (Directory.Exists(folder))
                    Directory.Delete(folder, true);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            { }
        }

        private static string GetCurrentDir()
        {
            var directoryInfo = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
            if (directoryInfo != null)
                return directoryInfo.FullName + Path.DirectorySeparatorChar;
            throw new DirectoryNotFoundException();
        }

        [TestCleanup]
        public void CleanUp()
        {
            foreach (var tempFolder in _tempFolders)
                DeleteTemporaryFolder(tempFolder);
        }
        #endregion
    }
}
