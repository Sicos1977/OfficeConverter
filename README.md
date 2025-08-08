OfficeConverter
===============

Convert Microsoft Office document to PDF (Office 2007 - 2016)

  - Converts .DOC, .DOCM, .DOCX, .DOT, .DOTM, .ODT, .XLS, .XLSB, .XLSM, .XLSX, .XLT, .XLTM, .XLTX, .XLW, .ODS, .POT, .PPT, .POTM, .POTX, .PPS, .PPSM, .PPSX, .PPTM, .PPTX and .ODP files to PDF with the use of Microsoft Office
  - Checks if the files are password protected without using Microsoft Office to speed up conversion

## License Information

OfficeConverter is Copyright (C)2014-2025 Kees van Spelde and is licensed under the MIT license:

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.

## Installing via NuGet

The easiest way to install OfficeConverter is via NuGet.

In Visual Studio's Package Manager Console, simply enter the following command:

    Install-Package OfficeConverter 

### Converting a file from code

```csharp
using (var converter = new Converter())
{
    converter.Convert(<inputfile>, <outputfile>);
}
```

## Logging

From version 1.7.0 OfficeConverter uses the Microsoft ILogger interface (https://docs.microsoft.com/en-us/dotnet/api/microsoft.extensions.logging.ilogger?view=dotnet-plat-ext-5.0). You can use any logging library that uses this interface.

OfficeConverter has some build in loggers that can be found in the ```OfficeConverter.Logger``` namespace. 

For example

```csharp
var logger = !string.IsNullOrWhiteSpace(<some logfile>)
                ? new OfficeConverter.Loggers.Stream(File.OpenWrite(<some logfile>))
                : new OfficeConverter.Loggers.Console();
```

Core Team
=========
    Sicos1977 (Kees van Spelde)

Support
=======
If you like my work then please consider a donation as a thank you.

## Reporting Bugs

Have a bug or a feature request? [Please open a new issue](https://github.com/Sicos1977/OfficeConverter/issues).

Before opening a new issue, please search for existing issues to avoid submitting duplicates.