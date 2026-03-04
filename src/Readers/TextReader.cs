using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Reads text content from Word 97-2003 binary documents.
/// Handles the CLX structure and Piece Table for both simple and complex documents.
/// 
/// Text in a .doc file is stored as a sequence of character positions (CPs).
/// For complex documents, the CLX structure in the Table stream contains
/// a Piece Table that maps CP ranges to file offsets in the WordDocument stream.
/// </summary>
public class TextReader
{
    private readonly BinaryReader _wordDocReader;
    private readonly FibReader _fib;
    private string _text = string.Empty;
    private List<Piece> _pieces = new();

    /// <param name="wordDocReader">Reader for the WordDocument stream</param>
    /// <param name="tableReader">Reader for the Table stream (0Table or 1Table)</param>
    /// <param name="fib">Parsed FIB</param>
    public TextReader(BinaryReader wordDocReader, BinaryReader tableReader, FibReader fib)
    {
        _wordDocReader = wordDocReader;
        TableReader = tableReader;
        _fib = fib;
    }

    /// <summary>
    /// Reader for the Table stream
    /// </summary>
    public BinaryReader TableReader { get; }

    /// <summary>
    /// Gets the complete reconstructed text of the main document body.
    /// </summary>
    public string Text => _text;

    /// <summary>
    /// Gets the piece table entries.
    /// </summary>
    public IReadOnlyList<Piece> Pieces => _pieces;

    /// <summary>
    /// Reads the complete text content from the main document body.
    /// </summary>
    public string ReadText()
    {
        if (_fib.CcpText <= 0) return string.Empty;

        if (_fib.FComplex)
        {
            // Complex document: parse CLX from Table stream to get piece table
            ReadClx();
            _text = ReconstructTextFromPieces(_fib.CcpText + _fib.CcpFtn + _fib.CcpHdd + _fib.CcpAtn + _fib.CcpEdn + _fib.CcpTxbx + _fib.CcpHdrTxbx);
        }
        else
        {
            // Simple (non-complex) document: text starts at offset 0 in WordDocument stream
            // after the FIB header. For Word 97+, text is at fcMin.
            // In practice, for non-complex docs, the text for CPs [0..ccpText) is at
            // a fixed offset. We use piece table logic uniformly when possible.
            // Fallback: read directly from WordDocument stream.
            ReadClx(); // Even non-complex docs may have a CLX
            if (_pieces.Count > 0)
            {
                _text = ReconstructTextFromPieces(_fib.CcpText + _fib.CcpFtn + _fib.CcpHdd + _fib.CcpAtn + _fib.CcpEdn + _fib.CcpTxbx + _fib.CcpHdrTxbx);
            }
            else
            {
                _text = ReadSimpleText();
            }
        }

        return _text;
    }

    /// <summary>
    /// Gets text for a specific CP range.
    /// </summary>
    public string GetText(int startCp, int length)
    {
        if (string.IsNullOrEmpty(_text) || startCp >= _text.Length)
            return string.Empty;

        var end = Math.Min(startCp + length, _text.Length);
        return _text.Substring(startCp, end - startCp);
    }

    // ─── CLX Parsing ────────────────────────────────────────────────

    /// <summary>
    /// Reads the CLX structure from the Table stream.
    /// CLX = array of Prc (optional grpprl prefixed by clxt=1) + Pcdt (piece table, clxt=2)
    /// </summary>
    private void ReadClx()
    {
        if (_fib.FcClx == 0 || _fib.LcbClx == 0)
            return;

        TableReader.BaseStream.Seek(_fib.FcClx, SeekOrigin.Begin);
        var endPosition = _fib.FcClx + _fib.LcbClx;

        // Skip any Prc entries (clxt = 0x01)
        while (TableReader.BaseStream.Position < endPosition)
        {
            var clxt = TableReader.ReadByte();

            if (clxt == 0x01)
            {
                // Prc — contains a GrpPrl
                var cbGrpprl = TableReader.ReadInt16();
                if (cbGrpprl > 0)
                    TableReader.BaseStream.Seek(cbGrpprl, SeekOrigin.Current);
            }
            else if (clxt == 0x02)
            {
                // Pcdt — the piece table
                var lcb = TableReader.ReadInt32(); // size of PlcPcd
                ReadPlcPcd(lcb);
                break;
            }
            else
            {
                // Unknown clxt — stop
                break;
            }
        }
    }

    /// <summary>
    /// Reads the PlcPcd (Piece Table) structure.
    /// 
    /// PlcPcd layout:
    ///   CP[0] CP[1] ... CP[n]       —  (n+1) × 4-byte CPs
    ///   PCD[0] PCD[1] ... PCD[n-1]  —  n × 8-byte Piece Descriptors
    ///
    /// where n = number of pieces.
    /// Total size = (n+1)*4 + n*8 = 4 + n*12
    /// So n = (lcb - 4) / 12
    /// </summary>
    private void ReadPlcPcd(int lcb)
    {
        if (lcb < 16) return; // minimum: 2 CPs + 1 PCD = 4+4+8 = 16

        var pieceCount = (lcb - 4) / 12;
        if (pieceCount <= 0) return;

        // Read CP array: (pieceCount + 1) entries
        var cps = new int[pieceCount + 1];
        for (int i = 0; i <= pieceCount; i++)
        {
            cps[i] = TableReader.ReadInt32();
        }

        // Read PCD array: pieceCount entries, each 8 bytes
        _pieces = new List<Piece>(pieceCount);
        for (int i = 0; i < pieceCount; i++)
        {
            var pcd = ReadPcd();
            var piece = new Piece
            {
                CpStart = cps[i],
                CpEnd = cps[i + 1],
                FileOffset = pcd.fc,
                IsUnicode = !pcd.fCompressed,
                Prm = pcd.prm
            };
            _pieces.Add(piece);
        }
    }

    /// <summary>
    /// Reads a single PCD (Piece Descriptor), 8 bytes:
    ///   ABCDxxxxh  (2 bytes) - first word, unused in practice
    ///   fc         (4 bytes) - file offset in WordDocument stream
    ///   prm        (2 bytes) - property modifier
    ///
    /// fc encoding:
    ///   If bit 30 is set (0x40000000), the text is ANSI (compressed).
    ///   The actual byte offset = (fc &amp; ~0x40000000) / 2  (for compressed)
    ///   or = fc  (for Unicode).
    /// </summary>
    private (uint fc, bool fCompressed, ushort prm) ReadPcd()
    {
        // Bytes 0-1: first word (ignored)
        TableReader.ReadUInt16();

        // Bytes 2-5: fc with encoding flag
        var rawFc = TableReader.ReadUInt32();

        bool fCompressed = (rawFc & 0x40000000) != 0;
        uint fc;
        if (fCompressed)
        {
            // ANSI text: real byte offset = (rawFc & ~0x40000000) / 2
            fc = (rawFc & 0x3FFFFFFF) / 2;
        }
        else
        {
            // Unicode text: fc is byte offset as-is
            fc = rawFc;
        }

        // Bytes 6-7: prm (property modifier)
        var prm = TableReader.ReadUInt16();

        return (fc, fCompressed, prm);
    }

    // ─── Text Reconstruction ────────────────────────────────────────

    /// <summary>
    /// Reconstructs the main document text from piece table entries.
    /// Includes all characters in the total CP range.
    /// </summary>
    private string ReconstructTextFromPieces(int totalCpCount)
    {
        if (_pieces.Count == 0) return string.Empty;

        // Use a safe total count for the builder
        var sb = new StringBuilder(Math.Max(0, totalCpCount));

        foreach (var piece in _pieces)
        {
            // Read all characters within the piece that fall within the requested total CP range
            var cpStart = piece.CpStart;
            var cpEnd = Math.Min(piece.CpEnd, totalCpCount);
            if (cpStart >= cpEnd) continue;
            if (cpStart >= totalCpCount) break;

            var charCount = cpEnd - cpStart;

            if (piece.IsUnicode)
            {
                // Unicode (UTF-16LE): each character is 2 bytes
                var byteOffset = piece.FileOffset;
                _wordDocReader.BaseStream.Seek(byteOffset, SeekOrigin.Begin);
                var bytes = _wordDocReader.ReadBytes(charCount * 2);
                sb.Append(Encoding.Unicode.GetString(bytes, 0, Math.Min(bytes.Length, charCount * 2)));
            }
            else
            {
                // ANSI (compressed): each character is 1 byte
                var byteOffset = piece.FileOffset;
                _wordDocReader.BaseStream.Seek(byteOffset, SeekOrigin.Begin);
                var bytes = _wordDocReader.ReadBytes(charCount);
                // 8-bit text in Word is usually ISO-8859-1 (Latin-1) or mapped directly to Unicode 0-255
                var encoding = Encoding.GetEncoding("iso-8859-1");
                sb.Append(encoding.GetString(bytes, 0, Math.Min(bytes.Length, charCount)));
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Fallback: reads text directly from WordDocument stream for non-complex documents
    /// that have no CLX. This is rare for Word 97+ but handled for safety.
    /// </summary>
    private string ReadSimpleText()
    {
        // For non-complex Word 97+ docs, text starts at CP 0 in the WordDocument stream
        // The byte offset depends on whether the text is Unicode or ANSI.
        // According to MS-DOC, for non-complex documents, text starts at fcMin
        // which is typically at offset 0x200 (512 bytes) after the FIB.
        // However, we should calculate this properly based on FIB size.

        var ccpText = _fib.CcpText;
        if (ccpText <= 0) return string.Empty;

        // Calculate the text offset based on FIB version
        // Word 97+ FIB is typically 512 bytes (0x200) for the base structure
        // But we should be more flexible and try different offsets
        var textOffset = 0x200; // Default offset for Word 97+

        // Try reading as Unicode from calculated offset
        try
        {
            _wordDocReader.BaseStream.Seek(textOffset, SeekOrigin.Begin);
            var bytes = _wordDocReader.ReadBytes(ccpText * 2);
            
            // Validate the text - check if it looks like valid Unicode
            var text = Encoding.Unicode.GetString(bytes);
            
            // If text contains too many null characters or control characters,
            // it might be ANSI or the offset is wrong
            var nullCount = text.Count(c => c == '\0');
            if (nullCount > ccpText * 0.5) // More than 50% nulls
            {
                // Try reading as ANSI
                _wordDocReader.BaseStream.Seek(textOffset, SeekOrigin.Begin);
                var ansiBytes = _wordDocReader.ReadBytes(ccpText);
                var ansiEncoding = Encoding.GetEncoding(1252);
                text = ansiEncoding.GetString(ansiBytes);
            }
            
            return text;
        }
        catch
        {
            return string.Empty;
        }
    }
}

/// <summary>
/// Represents a piece in the Piece Table.
/// Each piece maps a range of Character Positions (CPs) to a byte offset
/// in the WordDocument stream.
/// </summary>
public class Piece
{
    /// <summary>Starting CP (inclusive)</summary>
    public int CpStart { get; set; }

    /// <summary>Ending CP (exclusive)</summary>
    public int CpEnd { get; set; }

    /// <summary>Byte offset in the WordDocument stream</summary>
    public uint FileOffset { get; set; }

    /// <summary>True if text at this offset is Unicode (UTF-16LE), false if ANSI</summary>
    public bool IsUnicode { get; set; }

    /// <summary>Property modifier (Prm) from the PCD</summary>
    public ushort Prm { get; set; }

    /// <summary>Number of characters in this piece</summary>
    public int CharCount => CpEnd - CpStart;

    // Legacy compatibility properties
    public int Start { get => CpStart; set => CpStart = value; }
    public int End { get => CpEnd; set => CpEnd = value; }
    public int Length => CharCount;
}

/// <summary>
/// Sprm (Single Property Modifier) parser.
/// <summary>
/// Document Properties (DOP) Reader.
/// Reads from the Table stream at fcDop offset.
/// </summary>
public class DocumentPropertiesReader
{
    private readonly BinaryReader TableReader;
    private readonly FibReader _fib;

    public DocumentPropertiesReader(BinaryReader tableReader, FibReader fib)
    {
        TableReader = tableReader;
        _fib = fib;
    }

    /// <summary>
    /// Reads document properties from the Table stream.
    /// </summary>
    public DocumentProperties Read()
    {
        var props = new DocumentProperties();

        if (_fib.FcDop == 0 || _fib.LcbDop == 0) return props;

        TableReader.BaseStream.Seek(_fib.FcDop, SeekOrigin.Begin);

        // DOP is a variable-length structure. We read known fields.
        // Per MS-DOC §2.7.4 (Dop97):
        // The DOP structure has evolved across Word versions.
        // We read the common fields carefully.

        var dopBytes = TableReader.ReadBytes((int)Math.Min(_fib.LcbDop, 500));
        if (dopBytes.Length < 20) return props;

        // Offsets within DOP (Dop97 structure):
        // 0-1: bit flags (fWidowControl, fPaginated, fFacingPages, etc.)
        // 2-3: bit flags continued
        // 4-5: bit flags continued
        // 6-7: bit flags continued
        // 14-15: dxaTab (default tab width)
        // 16-17: dxaColumns (column width)
        // 28-29: itxtWrap (text wrapping)
        // For page setup, we rely on section properties (SEP) instead of DOP.
        // DOP primarily stores document-level flags.

        // Extract bit flags from Dop97 structure
        if (dopBytes.Length >= 4)
        {
            var flags0 = BitConverter.ToUInt16(dopBytes, 0);
            var flags1 = BitConverter.ToUInt16(dopBytes, 2);

            // Group 0 flags (from MS-DOC §2.7.4)
            props.FWidowControl = (flags0 & 0x0002) != 0;      // fWidowControl
            props.FPaginated = (flags0 & 0x0004) != 0;         // fPaginated
            props.FFacingPages = (flags0 & 0x0008) != 0;       // fFacingPages
            props.FBreaks = (flags0 & 0x0010) != 0;            // fBreaks
            props.FAutoHyphenate = (flags0 & 0x0020) != 0;     // fAutoHyphenate
            props.FDoHyphenation = (flags0 & 0x0040) != 0;     // fDoHyphenation
            props.FFELayout = (flags0 & 0x0080) != 0;          // fFELayout
            props.FLayoutSameAsWin95 = (flags0 & 0x0100) != 0; // fLayoutSameAsWin95
            props.FPrintBodyBeforeHeaders = (flags0 & 0x0200) != 0; // fPrintBodyBeforeHeaders
            props.FSuppressBottomSpacing = (flags0 & 0x0400) != 0; // fSuppressBottomSpacing
            props.FWrapAuto = (flags0 & 0x0800) != 0;          // fWrapAuto
            props.FPrintPaperBefore = (flags0 & 0x1000) != 0;  // fPrintPaperBefore
            props.FSuppressSpacings = (flags0 & 0x2000) != 0;  // fSuppressSpacings
            props.FMirrorMargins = (flags0 & 0x4000) != 0;     // fMirrorMargins
            // bits 14-15: fRuler, fNoTabForInd

            // Group 1 flags
            props.FUsePrinterMetrics = (flags1 & 0x0001) != 0; // fUsePrinterMetrics
            props.FNoPgp = (flags1 & 0x0002) != 0;             // fNoPgp
            props.FShrinkToFit = (flags1 & 0x0004) != 0;       // fShrinkToFit
            props.FPrintFormsData = (flags1 & 0x0008) != 0;    // fPrintFormsData
            props.FAllowPositionOnOnly = (flags1 & 0x0010) != 0; // fAllowPositionOnOnly
            props.FDisplayBackground = (flags1 & 0x0020) != 0; // fDisplayBackground
            props.FDisplayLineNumbers = (flags1 & 0x0040) != 0; // fDisplayLineNumbers
            props.FPrintMicros = (flags1 & 0x0080) != 0;       // fPrintMicros
            props.FSaveFormsData = (flags1 & 0x0100) != 0;     // fSaveFormsData
            props.FDisplayColBreak = (flags1 & 0x0200) != 0;   // fDisplayColBreak
            props.FDisplayPageEnd = (flags1 & 0x0400) != 0;    // fDisplayPageEnd
            props.FDisplayUnits = (flags1 & 0x0800) != 0;      // fDisplayUnits
            props.FProtectForms = (flags1 & 0x1000) != 0;      // fProtectForms
            props.FProtectSparce = (flags1 & 0x2000) != 0;     // fProtectSparce
            props.FConsecutiveHyphen = (flags1 & 0x4000) != 0; // fConsecutiveHyphen
            props.FLetterFinal = (flags1 & 0x8000) != 0;       // fLetterFinal
        }

        if (dopBytes.Length >= 16)
        {
            // dxaTab (default tab width) at offset 14-15
            props.DxaTab = BitConverter.ToInt16(dopBytes, 14);
            // dxaColumns (column width) at offset 16-17
            props.DxaColumns = BitConverter.ToInt16(dopBytes, 16);
        }

        if (dopBytes.Length >= 30)
        {
            // itxtWrap at offset 28-29
            props.ITxtWrap = BitConverter.ToInt16(dopBytes, 28);
        }

        // Default margins and page size will come from section properties
        // For now, return defaults that are reasonable
        props.PageWidth = 12240;   // 8.5" in twips
        props.PageHeight = 15840;  // 11" in twips
        props.MarginTop = 1440;    // 1" in twips
        props.MarginBottom = 1440;
        props.MarginLeft = 1800;   // 1.25" in twips
        props.MarginRight = 1800;

        return props;
    }
}
