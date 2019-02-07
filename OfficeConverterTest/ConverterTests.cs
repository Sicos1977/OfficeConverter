using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeConverter;
using OfficeConverter.Exceptions;

//
// ConverterTests.cs
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

namespace OfficeConverterTest
{
    [TestClass]
    public class ConverterTests
    {
        private readonly List<string> _tempFolders = new List<string>();

        [TestMethod]
        [ExpectedException(typeof(OCFileTypeNotSupported))]
        public void FileTypeNotSupported()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\filetypenotsupported.txt", outputFile);
        }

        [TestMethod]
        [ExpectedException(typeof(PasswordProtectedChecker.Exceptions.PPCFileIsCorrupt))]
        public void FileIsCorrupt()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A corrupt compound document.doc", outputFile);
        }

        #region Microsoft Office Word tests
        [TestMethod]
        public void DocWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A DOC word document without embedded files.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocWith7EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A DOC word document with 7 embedded files.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void DocWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A DOC word document with password.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using (var logStream = new MemoryStream())
            {
                using (var converter = new Converter(logStream))
                {
                    converter.Convert(GetCurrentDir() + "TestFiles\\A DOCX word document without embedded files.docx",
                        outputFile);
                }

                var log = Encoding.ASCII.GetString(logStream.ToArray());
                Assert.IsTrue(log.Contains("Document exported to PDF"));
            }

            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocxWith7EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using (var converter = new Converter())
            {
                using (var logStream = new MemoryStream())
                {
                    converter.Convert(GetCurrentDir() + "TestFiles\\A DOCX word document with 7 embedded files.docx",
                        outputFile, logStream);
                    var log = Encoding.ASCII.GetString(logStream.ToArray());
                    Assert.IsTrue(log.Contains("Document exported to PDF"));

                }
                using (var logStream = new MemoryStream())
                {
                    converter.Convert(GetCurrentDir() + "TestFiles\\A DOCX word document with 7 embedded files.docx",
                        outputFile, logStream);
                    var log = Encoding.ASCII.GetString(logStream.ToArray());
                    Assert.IsTrue(log.Contains("Document exported to PDF"));
                }
            }

            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocxWith7EmbeddedFiles10Times()
        {
            var currentDir = GetCurrentDir();

            using (var converter = new Converter())
            {
                for (var i = 0; i < 10; i++)
                {
                    var outputFile = CreateTemporaryFolder() + "\\test.pdf";
                    converter.Convert(currentDir + "TestFiles\\A DOCX word document with 7 embedded files.docx", outputFile);
                    Assert.IsTrue(File.Exists(outputFile));
                }
            }
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void DocxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A DOCX word document with password.docx", outputFile);
        }
        #endregion

        #region Microsoft Office Excel tests
        [TestMethod]
        public void XlsWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A XLS excel document without embedded files.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsWith2EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A XLS excel document with 2 embedded files.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsWith2EmbeddedFiles10Times()
        {
            var currentDir = GetCurrentDir();

            using (var converter = new Converter())
            {
                for (var i = 0; i < 10; i++)
                {
                    var outputFile = CreateTemporaryFolder() + "\\test.pdf";
                    converter.Convert(currentDir + "TestFiles\\A XLS excel document with 2 embedded files.xls", outputFile);
                    Assert.IsTrue(File.Exists(outputFile));
                }
            }
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void XlsWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A XLS excel document with password.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A XLSX excel document without embedded files.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsxWith2EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A XLSX excel document with 2 embedded files.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void XlsxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A XLSX excel document with password.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvSemicolonSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\Semicolon separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvCommaSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\Comma separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvSpaceSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\Space separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void CsvTabSeparated()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\Tab separated csv.csv", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }
        #endregion

        #region Microsoft Office PowerPoint tests
        [TestMethod]
        public void PptWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A PPT PowerPoint document without embedded files.ppt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void PptWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A PPT powerpoint document with 3 embedded files.ppt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void PptWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A PPT PowerPoint document with password.ppt", outputFile);
        }

        [TestMethod]
        public void PptxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A PPTX PowerPoint document without embedded files.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void PptxWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A PPTX powerpoint document with 3 embedded files.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void PptxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\A PPTX PowerPoint document with password.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }
        #endregion

        #region Open Office Writer tests
        [TestMethod]
        public void OdtWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\An ODT document without embedded files.odt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void OdtWith8EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\An ODT document with 8 embedded files.odt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void OdtWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\An ODT document with password.odt", outputFile);
        }
        #endregion
        
        #region Open Office Impress tests
        [TestMethod]
        public void OdpWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\An ODP document without embedded files.odp", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void OdpWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\An ODP document with 3 embedded files.odp", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void OdpWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            using(var converter = new Converter())
                converter.Convert(GetCurrentDir() + "TestFiles\\An ODP document with password.odp", outputFile);
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
