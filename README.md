# Nedev.FileConverters.DocToDocx

A DOC to DOCX converter for .NET 8.0 and .NET Standard 2.1 with no third-party runtime dependencies.

## Features

- **Binary DOC reader**: Reads compound file storage, FIB metadata, CLX/piece tables, FKPs, PLCFs, and related legacy Word structures.
- **Rich text and paragraph formatting**: Preserves common character and paragraph properties including fonts, size, bold, italic, underline, colors, highlight, alignment, spacing, indentation, borders, shading, and list formatting.
- **Tables**: Writes table width, indentation, spacing, borders, shading, header rows, cantSplit, vertical merges, horizontal merges, and nested table structures recovered from TAP/PAP data.
- **Sections and page setup**: Emits page size, orientation, margins, page numbering, and first/odd/even header and footer parts.
- **Images and floating pictures**: Extracts embedded images and OfficeArt picture data, writes media parts, and emits inline or anchored drawings depending on available FSPA anchor information.
- **Footnotes, endnotes, comments, textboxes, and equations**: Writes common annotation parts, textbox content, and Equation Editor content converted to OMML where recognized.
- **Encrypted DOC support**: Supports XOR-obfuscated streams and Office 97-2003 RC4-encrypted DOC files when the correct password is supplied.
- **Package validation helpers**: Exposes helpers to load, save, convert, and validate generated DOCX packages.
- **CLI and library APIs**: Includes synchronous conversion, asynchronous conversion, progress reporting, and a command-line tool for file and directory conversion.

## Current limitations and known issues

The converter is intentionally pragmatic: it aims to produce valid, readable DOCX output from a wide range of legacy DOC files, but it does not implement the full MS-DOC / OfficeArt feature surface.

- **Complex vector shapes are downgraded**: Non-picture OfficeArt content and SmartArt-like shapes are currently written through simplified DrawingML fallback paths rather than full-fidelity shape reconstruction.
- **Chart support is partial**: Embedded BIFF/OLE chart data is scanned on a best-effort basis. Category and value series can often be recovered, but chart formatting, axis details, legends, and some chart types are regenerated with minimal defaults.
- **Theme parsing is basic**: Theme XML can be extracted from legacy storage, but theme color/style interpretation is not fully implemented.
- **Best-effort parsing may skip malformed content**: Some OLE, chart, image, and binary parsing paths intentionally catch truncated or malformed input and continue conversion instead of failing the entire document.
- **Layout heuristics remain in a few areas**: Deeply nested tables, unusual merge layouts, and some header/footer or drawing placement cases may be approximated rather than reproduced exactly.
- **Not a round-trip converter**: The goal is compatible DOCX output, not byte-for-byte structural equivalence with the source DOC.

## Priority implementation gaps

The following areas are the main remaining implementation and maintenance hotspots identified from the current codebase.

1. **RC4 stream boundary handling needs a dedicated end-to-end audit**: The RC4 setup in the DOC reader includes explicit uncertainty around where encrypted content begins in WordDocument and related streams. The current code works for covered cases, but this area should be treated as fragile until it is backed by encrypted real-world regression samples.
2. **Silent best-effort recovery can hide data loss**: Several parsing paths intentionally swallow malformed or truncated content and continue conversion. That keeps output generation resilient, but it also means missing charts, images, anchors, or OLE payloads can fail quietly instead of producing structured diagnostics.
3. **Chart reconstruction remains intentionally minimal**: The BIFF scanner can recover simple category/value grids, but chart metadata and formatting are still regenerated from defaults. This is sufficient for editable fallback charts, not faithful reproduction of complex embedded Excel chart state.
4. **Theme support stops at extraction**: Theme XML is stored and preserved when found, but theme color/style interpretation is still placeholder-level and does not drive broader formatting reconstruction.
5. **Nested table parsing is still a high-complexity area**: Table recovery relies on paragraph nesting levels, cell boundary markers, and stack-based state. The implementation is robust for common cases, but exotic nesting and malformed cell boundaries remain one of the more failure-prone parts of the reader.
6. **Binary property decoding still needs broader hostile-input coverage**: SPRM/FKP parsing has targeted regression tests, but there is still no broad malformed-input or fuzz-style test coverage for truncated operands, corrupt piece tables, or extreme legacy edge cases.

## Test coverage notes

- **Covered today**: Package validation, sample conversion flows, selected chart heuristics, OMML namespace generation, FIB regression cases, bookmark decoding, and some encryption helper behavior.
- **Still missing**: End-to-end encrypted DOC regression files, malformed document fuzzing, broad OLE/object-pool failure cases, and systematic compatibility suites for complex layout documents.

## Usage

### Library

Reference the Nedev.FileConverters.DocToDocx assembly and call the public converter API:

```csharp
using Nedev.FileConverters.DocToDocx;

DocToDocxConverter.Convert("input.doc", "output.docx");

DocToDocxConverter.Convert(
    "input.doc",
    "output.docx",
    password: null,
    enableHyperlinks: false);
```

### Async conversion with progress

```csharp
using Nedev.FileConverters.DocToDocx;

var progress = new Progress<ConversionProgress>(update =>
{
    Console.WriteLine($"[{update.PercentComplete,3}%] {update.Stage}: {update.Message}");
});

await DocToDocxConverter.ConvertAsync(
    "input.doc",
    "output.docx",
    progress,
    password: null,
    enableHyperlinks: true,
    cancellationToken: CancellationToken.None);
```

The converter can also detect an existing DOCX input and copy it through to the output path instead of attempting DOC parsing.

### Command-line interface

```bash
Nedev.FileConverters.DocToDocx.Cli <input.doc|input.docx|inputDir> <output.docx|outputDir> [-p <password>] [-r] [--no-hyperlinks]
```

When the input is a DOC file, the tool converts it. When the input is already a DOCX file, it copies the package to the output location. When the input is a directory, the tool can convert matching files in-place into a mirrored output directory tree.

**Arguments:**
- `<input.doc|input.docx|inputDir>` Input DOC file, DOCX file, or directory.
- `<output.docx|outputDir>` Output DOCX file or destination directory.

**Options:**
- `-p`, `--password` Password for encrypted DOC files.
- `-r`, `--recursive` Recursively process directories.
- `--no-hyperlinks` Disable hyperlink elements and relationships in the generated DOCX.
- `-v`, `--version` Print the CLI version and exit.
- `-h`, `--help` Show usage information and exit.

## Running the tests

The repository includes unit and focused integration tests under `src/Nedev.FileConverters.DocToDocx.Tests`.

```bash
dotnet test
```

The current test suite covers package validation helpers, selected writer behavior, chart scanning heuristics, OMML namespace generation, sample document conversion, and CLI conversion flows. It does not yet provide exhaustive coverage for malformed-input fuzzing or every edge case in legacy Word binary parsing.

A GitHub Actions workflow at `.github/workflows/ci.yml` builds the solution and runs the test suite on push and pull request events.
