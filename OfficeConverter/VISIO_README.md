# Visio Converter Support

This project includes support for converting Microsoft Visio files to PDF format.

## Enabling Visio Support

To enable Visio conversion functionality, you need to:

1. **Add the Visio Interop Assembly**
   - Obtain the `Microsoft.Office.Interop.Visio.dll` file from your Microsoft Office installation or from the Primary Interop Assemblies (PIA)
   - Copy the DLL to the `OfficeAssemblies` folder in this project

2. **Update the Project File**
   - Open `OfficeConverter.csproj`
   - Uncomment the Visio reference section (around line 85-89):
     ```xml
     <Reference Include="Microsoft.Office.Interop.Visio">
       <HintPath>OfficeAssemblies\Microsoft.Office.Interop.Visio.dll</HintPath>
     </Reference>
     ```

3. **Enable Conditional Compilation**
   - Add the `VISIO_INTEROP` compilation symbol to your project
   - In Visual Studio: Project Properties → Build → Conditional compilation symbols
   - Or add to the `.csproj` file:
     ```xml
     <PropertyGroup>
       <DefineConstants>$(DefineConstants);VISIO_INTEROP</DefineConstants>
     </PropertyGroup>
     ```

## Supported Visio File Formats

Once enabled, the converter supports the following Visio file extensions:
- `.VSD` - Visio Drawing (2003-2010)
- `.VSDX` - Visio Drawing (2013+)
- `.VDX` - Visio XML Drawing
- `.VSS` - Visio Stencil
- `.VSSX` - Visio Stencil (2013+)
- `.VST` - Visio Template
- `.VSTX` - Visio Template (2013+)
- `.VDW` - Visio Web Drawing

## LibreOffice Support

Visio files can also be converted using LibreOffice Draw by setting the `UseLibreOffice` property to `true` on the `Converter` object. This does not require the Visio Interop assembly.

## Requirements

- Microsoft Visio 2003 or later must be installed on the system
- The application must have permission to automate Visio

## Credits

Initial implementation by Heinrich Elsigan, based on the PowerPoint converter pattern.
