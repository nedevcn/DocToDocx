# Nedev.DocToDocx

A high‑fidelity `.doc` → `.docx` converter for .NET 10 with no third‑party dependencies.

## Features

- **Binary `.doc` reader**: Implements core MS‑DOC structures (CFB, FIB, CLX/Piece Table, CHPX/PAPX FKPs, PLCFs).
- **Rich text & styles**: Fonts, font sizes, bold/italic/underline, colors, highlighting, outline/emboss/shadow, language (`w:lang`), paragraph alignment, spacing, indentation, borders, shading, and numbered/bulleted lists (including many localized formats).
- **Tables**: Multi‑row/column tables with row height, header rows, cantSplit, alignment, per‑cell width, and TAP‑driven layout; writer side supports proper vertical merges (`vMerge restart/continue`).
- **Sections & page setup**: Multiple sections with page size/orientation, margins, starting page number, and First/Odd/Even headers/footers mapped to separate DOCX parts.
- **Images**: Extracts embedded images from the `Data` stream (PNG/JPEG/GIF/BMP/OfficeArt BLIPs), writes `word/media/*`, generates `w:drawing` with size inferred from image dimensions and auto‑scaled to page width, plus basic alt text.
- **Footnotes, endnotes, comments, textboxes**: Reads and writes common note and annotation structures into DOCX footnotes/endnotes parts and DrawingML textboxes.
- **Encryption (XOR)**: Supports Word’s XOR‑obfuscated streams via `EncryptionHelper` and decrypted CFB streams.
- **No external dependencies**: Pure .NET, streaming writers (`XmlWriter`) for high performance and low memory usage.

> Note: While many MS‑DOC features are implemented, the converter does not yet claim 100% coverage of the full [MS‑DOC] specification. Complex OfficeArt shapes, OLE objects, and some rare formatting cases are intentionally out of scope for now.

## Library usage

Add a reference to the `Nedev.DocToDocx` assembly and call the static converter API:

```csharp
using Nedev.DocToDocx;

DocToDocxConverter.Convert("input.doc", "output.docx");
```

> This repository currently exposes the converter as a **library API only**.  
> If you need a CLI or additional validation tooling, you can build it on top of `DocToDocxConverter` according to your own application’s needs.