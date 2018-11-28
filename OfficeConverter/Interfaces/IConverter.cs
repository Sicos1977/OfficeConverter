using System.Runtime.InteropServices;

//
// IConverter.cs
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

namespace OfficeConverter.Interfaces
{
    /// <summary>
    ///     Interface to make Reader class COM exposable
    /// </summary>
    public interface IConverter
    {
        /// <summary>
        ///     Converts the <paramref name="inputFile" /> to PDF and saves it as the <paramref name="outputFile" />
        /// </summary>
        /// <param name="inputFile">The Microsoft Office file</param>
        /// <param name="outputFile">The output file with full path</param>
        /// <param name="useLibreOffice">
        ///     When set to <c>true</c> then LibreOffice is used to convert the file to PDF instead of
        ///     Microsoft Office
        /// </param>
        /// <returns>
        ///     Returns true when the conversion is succesfull, false is retournerd when an exception occurred.
        ///     The exception can be retrieved with the <see cref="GetErrorMessage" /> method
        /// </returns>
        [DispId(1)]
        bool ConvertFromCom(string inputFile, string outputFile, bool useLibreOffice);

        /// <summary>
        ///     Get the last know error message. When the string is empty there are no errors
        /// </summary>
        /// <returns></returns>
        [DispId(2)]
        string GetErrorMessage();
    }
}