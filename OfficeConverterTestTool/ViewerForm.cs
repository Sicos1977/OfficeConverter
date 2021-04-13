using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeConverter;

//
// ViewerForm.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2014-2021 Magic-Sessions. (www.magic-sessions.com)
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

namespace OfficeConverterTestTool
{
    public partial class ViewerForm : Form
    {
        readonly List<string> _tempFolders = new List<string>();

        #region ViewerForm
        public ViewerForm()
        {
            InitializeComponent();
        }
        #endregion

        #region ViewerForm_Load
        private void ViewerForm_Load(object sender, EventArgs e)
        {
            Closed += ViewerForm_Closed;
        }
        #endregion

        #region ViewerForm_Closed
        private void ViewerForm_Closed(object sender, EventArgs e)
        {
            foreach (var tempFolder in _tempFolders)
            {
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
        }
        #endregion

        #region SelectButton_Click
        private void SelectButton_Click(object sender, EventArgs e)
        {
            // Create an instance of the opeKn file dialog box.
            var openFileDialog1 = new OpenFileDialog
            {
                // ReSharper disable once LocalizableElement
                Filter = "Microsoft Office files|*.DOC;*.DOT;*.DOCM;*.DOCX;*.DOTM;*.ODT;*.XML;*.RTF;*.MHT;" +
                         "*.WPS;*.WRI;*.XLS;*.XLT;*.XLW;*.XLSB;*.XLSM;*.XLSX;" +
                         "*.XLTM;*.XLTX;*.CSV;*.ODS;*.POT;*.PPT;*.PPS;*.POTM;" +
                         "*.POTX;*.PPSM;*.PPSX;*.PPTM;*.PPTX;*.ODP",
                FilterIndex = 1,
                Multiselect = false
            };

            // Process input if the user clicked OK.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Open the selected file to read.
                string tempFolder = null;

                try
                {
                    tempFolder = GetTemporaryFolder();
                    _tempFolders.Add(tempFolder);

                    var outputFile = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.LastIndexOf('.')) + ".pdf";

                    OutputTextBox.Clear();
                    OutputTextBox.Text = @"Converting...";

                    using (var memoryStream = new MemoryStream())
                    using (var converter = new Converter(memoryStream))
                    {
                        converter.UseLibreOffice = LibreOfficeCheckBox.Checked;
                        converter.Convert(openFileDialog1.FileName, outputFile);
                        OutputTextBox.Text += Environment.NewLine + Encoding.Default.GetString(memoryStream.ToArray());
                        OutputTextBox.Text += Environment.NewLine + @"Converted file written to '" + outputFile + @"'";
                    }
                }
                catch (Exception ex)
                {
                    if (tempFolder != null && Directory.Exists(tempFolder))
                        Directory.Delete(tempFolder, true);

                    OutputTextBox.Text = GetInnerException(ex);
                }
            }
        }
        #endregion

        #region GetTemporaryFolder
        public string GetTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
        #endregion

        #region GetInnerException
        /// <summary>
        /// Get the complete inner exception tree
        /// </summary>
        /// <param name="e">The exception object</param>
        /// <returns></returns>
        public static string GetInnerException(Exception e)
        {
            var exception = e.Message + Environment.NewLine;
            if (e.InnerException != null)
                exception += GetInnerException(e.InnerException);
            return exception;
        }
        #endregion
    }
}
