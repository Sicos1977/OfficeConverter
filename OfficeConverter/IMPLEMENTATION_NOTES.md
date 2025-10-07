# Implementation Notes: Visio Converter

## Overview
This implementation adds support for converting Microsoft Visio files to PDF format in the OfficeConverter library.

## Key Implementation Details

### 1. Conditional Compilation
The Visio converter is implemented using conditional compilation with the `VISIO_INTEROP` symbol. This allows the library to build and work without the Visio Interop assembly, making it optional.

**Why this approach?**
- Microsoft.Office.Interop.Visio.dll is not available via NuGet
- Users need to obtain it from their Office installation
- The library should still compile and work for Word, Excel, and PowerPoint without Visio support

### 2. Files Modified

#### Visio.cs (NEW)
- Complete Visio converter implementation
- Based on PowerPoint.cs pattern
- Includes version detection (2003-2016)
- Handles document opening, conversion, and cleanup
- Manages Visio process lifecycle

#### Converter.cs
- Added Visio field and property (conditionally compiled)
- Added handling for Visio file extensions: .VSD, .VSDX, .VDX, .VSS, .VSSX, .VST, .VSTX, .VDW
- Updated error message to include Visio file types
- Added Visio disposal in Dispose() method

#### LibreOffice.cs
- Added Visio file type support in GetFilterType() method
- Uses "draw_pdf_Export" filter for Visio files
- Allows LibreOffice to convert Visio files as an alternative to Microsoft Visio

#### OfficeConverter.csproj
- Added commented-out reference to Microsoft.Office.Interop.Visio.dll
- Includes instructions for enabling the reference

#### VISIO_README.md (NEW)
- Comprehensive setup instructions
- Lists all supported Visio file formats
- Explains how to enable Visio support
- Credits Heinrich Elsigan for initial contribution

#### README.md
- Updated to mention Visio support
- Links to VISIO_README.md for setup details

### 3. Supported Visio File Formats
- `.VSD` - Visio Drawing (2003-2010)
- `.VSDX` - Visio Drawing (2013+)
- `.VDX` - Visio XML Drawing
- `.VSS` - Visio Stencil
- `.VSSX` - Visio Stencil (2013+)
- `.VST` - Visio Template
- `.VSTX` - Visio Template (2013+)
- `.VDW` - Visio Web Drawing

### 4. Usage

#### With Microsoft Visio (requires setup):
```csharp
// After adding the Visio Interop DLL and enabling VISIO_INTEROP
using (var converter = new Converter())
{
    converter.Convert("diagram.vsdx", "output.pdf");
}
```

#### With LibreOffice (works immediately):
```csharp
using (var converter = new Converter())
{
    converter.UseLibreOffice = true;
    converter.Convert("diagram.vsdx", "output.pdf");
}
```

### 5. Enabling Full Visio Support

To enable Visio conversion with Microsoft Visio:

1. Locate `Microsoft.Office.Interop.Visio.dll` from your Office installation
2. Copy it to `OfficeConverter/OfficeAssemblies/`
3. Uncomment the reference in `OfficeConverter.csproj`
4. Add `VISIO_INTEROP` to compilation symbols

Without these steps, Visio files can still be converted using LibreOffice.

### 6. Design Consistency
The implementation follows the exact same pattern as the existing Word, Excel, and PowerPoint converters:
- Version detection via registry
- Process management
- Document lifecycle handling
- Resiliency key cleanup
- Error handling
- IDisposable implementation

## Credits
- Original contribution: Heinrich Elsigan
- Integration: OfficeConverter team
- Based on: PowerPoint converter pattern
