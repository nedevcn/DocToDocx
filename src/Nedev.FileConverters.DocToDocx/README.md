# Nedev.FileConverters.DocToDocx

A high-performance DOC to DOCX converter for .NET 8.0 and .NET Standard 2.1 without third-party dependencies.

## Features

- **Binary .doc reader**: Implements core MS-DOC structures (CFB, FIB, CLX/Piece Table, CHPX/PAPX FKPs, PLCFs)
- **Rich text & styles**: Fonts, font sizes, bold/italic/underline, colors, highlighting, outline/emboss/shadow, language, paragraph alignment, spacing, indentation, borders, shading, and numbered/bulleted lists
- **Tables**: Multi-row/column tables with TAP-driven layout, row height rules, header rows, cantSplit, per-cell width, proper vertical/horizontal merges, table-level borders and shading
- **Sections & page setup**: Multiple sections with page size/orientation, margins, starting page number, and First/Odd/Even headers/footers
- **Images**: Extracts embedded images (PNG/JPEG/GIF/BMP/OfficeArt BLIPs), generates w:drawing with size inferred from image dimensions
- **OfficeArt pictures & floating anchors**: Parses Escher/OfficeArt records and FSPA anchors to recover picture shapes
- **Footnotes, endnotes, comments, textboxes**: Reads and writes common note and annotation structures
- **Encryption**: Supports XOR-obfuscated streams and Office 97-2003 RC4-encrypted documents
- **No external dependencies**: Pure .NET, streaming writers (XmlWriter) for high performance and low memory usage

## Getting Started

### Installation

```bash
dotnet add package Nedev.FileConverters.DocToDocx
```

### Basic Usage

```csharp
using Nedev.FileConverters.DocToDocx;

// Basic conversion
DocToDocxConverter.Convert("input.doc", "output.docx");

// Conversion with password
DocToDocxConverter.Convert("input.doc", "output.docx", password: "mypassword");

// Conversion without hyperlinks
DocToDocxConverter.Convert("input.doc", "output.docx", enableHyperlinks: false);
```

### Using with Nedev.FileConverters.Core

This package integrates with [Nedev.FileConverters.Core](https://www.nuget.org/packages/Nedev.FileConverters.Core) and supports automatic converter discovery:

```csharp
using Nedev.FileConverters;

// Convert using the unified Core API
using var input = File.OpenRead("input.doc");
using var output = Converter.Convert(input, "doc", "docx");
using var outputStream = File.Create("output.docx");
await output.CopyToAsync(outputStream);
```

## Supported Frameworks

- .NET 8.0
- .NET Standard 2.1

## Dependencies

- Nedev.FileConverters.Core 0.1.0
- System.Text.Encoding.CodePages 8.0.0

## License

This project is licensed under the MIT License.

## Links

- [GitHub Repository](https://github.com/nedevcn/FileConverters.DocToDocx)
- [NuGet Package](https://www.nuget.org/packages/Nedev.FileConverters.DocToDocx)
- [Nedev.FileConverters.Core](https://www.nuget.org/packages/Nedev.FileConverters.Core)