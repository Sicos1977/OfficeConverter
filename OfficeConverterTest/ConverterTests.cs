using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeConverter;
using OfficeConverter.Exceptions;

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
            new Converter().Convert("TestFiles\\filetypenotsupported.txt", outputFile);
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsCorrupt))]
        public void FileIsCorrupt()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A corrupt compound document.doc", outputFile);
        }

        #region Microsoft Office Word tests
        [TestMethod]
        public void DocWithoutEmbeddedFiles()
        {
           var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A DOC word document without embedded files.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocWith7EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A DOC word document with 7 embedded files.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void DocWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A DOC word document with password.doc", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A DOCX word document without embedded files.docx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void DocxWith7EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A DOCX word document with 7 embedded files.docx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void DocxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A DOCX word document with password.docx", outputFile);
        }
        #endregion

        #region Microsoft Office Excel tests
        [TestMethod]
        public void XlsWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A XLS excel document without embedded files.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsWith2EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A XLS excel document with 2 embedded files.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void XlsWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A XLS excel document with password.xls", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A XLSX excel document without embedded files.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void XlsxWith2EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A XLSX excel document with 2 embedded files.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void XlsxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A XLSX excel document with password.xlsx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }
        #endregion

        #region Microsoft Office PowerPoint tests
        [TestMethod]
        public void PptWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A PPT PowerPoint document without embedded files.ppt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void PptWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A PPT powerpoint document with 3 embedded files.ppt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void PptWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A PPT PowerPoint document with password.ppt", outputFile);
        }

        [TestMethod]
        public void PptxWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A PPTX PowerPoint document without embedded files.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void PptxWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A PPTX powerpoint document with 3 embedded files.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void PptxWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\A PPTX PowerPoint document with password.pptx", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }
        #endregion

        #region Open Office Writer tests
        [TestMethod]
        public void OdtWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\An ODT document without embedded files.odt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void OdtWith8EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\An ODT document with 8 embedded files.odt", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void OdtWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\An ODT document with password.odt", outputFile);
        }
        #endregion
        
        #region Open Office Impress tests
        [TestMethod]
        public void OdpWithoutEmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\An ODP document without embedded files.odp", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        public void OdpWith3EmbeddedFiles()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\An ODP document with 3 embedded files.odp", outputFile);
            Assert.IsTrue(File.Exists(outputFile));
        }

        [TestMethod]
        [ExpectedException(typeof(OCFileIsPasswordProtected))]
        public void OdpWithPassword()
        {
            var outputFile = CreateTemporaryFolder() + "\\test.pdf";
            new Converter().Convert("TestFiles\\An ODP document with password.odp", outputFile);
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

        [TestCleanup]
        public void CleanUp()
        {
            foreach (var tempFolder in _tempFolders)
                DeleteTemporaryFolder(tempFolder);
        }
        #endregion
    }
}
