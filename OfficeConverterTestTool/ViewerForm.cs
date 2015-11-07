using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeConverter;

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
                Filter = "Microsoft Office files|*.DOC;*.DOT;*.DOCM;*.DOCX;*.DOTM;*.ODT;*.RTF;*.MHT;" +
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

                    var extractor = new Converter();
                    var outputFile = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.LastIndexOf('.')) +
                                     ".pdf";
                    OutputTextBox.Text = "Converting...";
                    extractor.Convert(openFileDialog1.FileName, outputFile);
                    OutputTextBox.Clear();
                    OutputTextBox.Text = "Converted file written to '" + outputFile + "'";
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
