using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// FKP (Formatted Disk Page) parser for Word 97-2003 binary format.
/// FKPs store character (CHP) and paragraph (PAP) formatting properties.
/// </summary>
public class FkpParser
{
    private readonly BinaryReader _wordDocReader;
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly SprmParser _sprmParser;
    private readonly TextReader _textReader;

    private readonly Dictionary<uint, ChpFkp> _chpFkpCache = new();
    private readonly Dictionary<uint, PapFkp> _papFkpCache = new();

    public FkpParser(BinaryReader wordDocReader, BinaryReader tableReader, FibReader fib, TextReader textReader)
    {
        _wordDocReader = wordDocReader;
        _tableReader = tableReader;
        _fib = fib;
        _textReader = textReader;
        _sprmParser = new SprmParser(wordDocReader, 0);
    }

    #region CHP (Character Properties)

    public Dictionary<int, ChpBase> ReadChpProperties()
    {
        var chpMap = new Dictionary<int, ChpBase>();
        if (_fib.FcPlcfBteChpx == 0 || _fib.LcbPlcfBteChpx < 16) return chpMap;

        _tableReader.BaseStream.Seek(_fib.FcPlcfBteChpx, SeekOrigin.Begin);
        
        // PLC structure: CP array (n+1 entries) + PCD array (n entries)
        // Each CP is 4 bytes, each PCD is 8 bytes
        // Total size = 4 + n*12, so n = (lcb - 4) / 12
        var lcb = (int)_fib.LcbPlcfBteChpx;
        var numPcd = (lcb - 4) / 12;
        if (numPcd <= 0) return chpMap;
        
        // Read CP array: numPcd + 1 entries
        var cpArray = new int[numPcd + 1];
        for (int i = 0; i <= numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            cpArray[i] = _tableReader.ReadInt32();
        }
        
        // Read PCD array: numPcd entries (PNs)
        var pnArray = new uint[numPcd];
        for (int i = 0; i < numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            pnArray[i] = _tableReader.ReadUInt32();
        }
        
        for (int i = 0; i < numPcd; i++)
        {
            var fkp = GetChpFkp(pnArray[i]);
            if (fkp == null) continue;
            
            foreach (var entry in fkp.Entries)
            {
                // FKP entries contain absolute File Character (FC) values
                int startCp = FcToCp(entry.StartCpOffset);
                int endCp = FcToCp(entry.EndCpOffset);
                
                int finalStart = Math.Max(0, startCp);
                int finalEnd = Math.Min(_fib.CcpText, endCp);
                
                for (int cp = finalStart; cp < finalEnd; cp++) chpMap[cp] = entry.Properties;
            }
        }
        return chpMap;
    }

    private ChpFkp? GetChpFkp(uint pn)
    {
        if (_chpFkpCache.TryGetValue(pn, out var cached)) return cached;
        var fkp = LoadChpFkp(pn);
        if (fkp != null) _chpFkpCache[pn] = fkp;
        return fkp;
    }

    private ChpFkp? LoadChpFkp(uint pn)
    {
        var offset = pn * WordConsts.FKP_PAGE_SIZE;
        if (offset + WordConsts.FKP_PAGE_SIZE > _wordDocReader.BaseStream.Length) return null;
        _wordDocReader.BaseStream.Seek(offset, SeekOrigin.Begin);
        return ParseChpFkp(_wordDocReader.ReadBytes(WordConsts.FKP_PAGE_SIZE));
    }

    private ChpFkp ParseChpFkp(byte[] data)
    {
        var fkp = new ChpFkp();
        
        // Safety check: data must be FKP_PAGE_SIZE bytes
        if (data.Length < WordConsts.FKP_PAGE_SIZE) return fkp;
        
        // Per MS-DOC spec: crun is the LAST byte of the 512-byte FKP page
        var crun = data[WordConsts.FKP_PAGE_SIZE - 1];
        if (crun == 0 || crun > 101) return fkp; // max crun for CHPX FKP is 101

        // FKP layout:
        //   rgfc[0..crun] : (crun+1) x 4-byte FC/CP values at offset 0
        //   rgb[0..crun-1] : crun x 1-byte offsets (word offsets into this page)
        //   ... property data ...
        //   crun : 1 byte at offset 511
        
        var rgfcSize = (crun + 1) * 4;
        if (rgfcSize + crun > data.Length) return fkp;

        // Read FC/CP array (crun+1 entries, starting at offset 0)
        var fcArray = new int[crun + 1];
        for (int i = 0; i <= crun; i++)
        {
            fcArray[i] = BitConverter.ToInt32(data, i * 4);
        }

        // Read property offset bytes (crun entries, starting after FC array)
        var rgbBase = rgfcSize;

        for (int i = 0; i < crun; i++)
        {
            if (rgbBase + i >= data.Length) break;
            
            var propOffset = data[rgbBase + i];
            if (propOffset == 0) continue;

            // propOffset is a word offset (multiply by 2 to get byte offset)
            var dataOffset = propOffset * 2;
            if (dataOffset >= WordConsts.FKP_PAGE_SIZE || dataOffset >= data.Length) continue;

            var cb = data[dataOffset];
            if (cb == 0 || dataOffset + 1 + cb > data.Length) continue;
            
            var chp = new ChpBase();
            var grpprl = new byte[cb];
            Array.Copy(data, dataOffset + 1, grpprl, 0, cb);
            _sprmParser.ApplyToChp(grpprl, chp);

            fkp.Entries.Add(new ChpFkpEntry
            {
                StartCpOffset = fcArray[i],
                EndCpOffset = fcArray[i + 1],
                Properties = chp
            });
        }
        return fkp;
    }

    #endregion

    #region PAP (Paragraph Properties)

    public Dictionary<int, PapBase> ReadPapProperties()
    {
        var papMap = new Dictionary<int, PapBase>();
        if (_fib.FcPlcfBtePapx == 0 || _fib.LcbPlcfBtePapx < 16) return papMap;

        _tableReader.BaseStream.Seek(_fib.FcPlcfBtePapx, SeekOrigin.Begin);
        
        // PLC structure: CP array (n+1 entries) + PCD array (n entries)
        // Each CP is 4 bytes, each PCD is 8 bytes
        // Total size = 4 + n*12, so n = (lcb - 4) / 12
        var lcb = (int)_fib.LcbPlcfBtePapx;
        var numPcd = (lcb - 4) / 12;
        if (numPcd <= 0) return papMap;
        
        // Read CP array: numPcd + 1 entries
        var cpArray = new int[numPcd + 1];
        for (int i = 0; i <= numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            cpArray[i] = _tableReader.ReadInt32();
        }
        
        // Read PCD array: numPcd entries (PNs)
        var pnArray = new uint[numPcd];
        for (int i = 0; i < numPcd; i++) 
        {
            if (_tableReader.BaseStream.Position + 4 > _tableReader.BaseStream.Length) break;
            pnArray[i] = _tableReader.ReadUInt32();
        }
        
        for (int i = 0; i < numPcd; i++)
        {
            var fkp = GetPapFkp(pnArray[i]);
            if (fkp == null) continue;
            
            foreach (var entry in fkp.Entries)
            {
                // FKP entries contain absolute File Character (FC) values
                int startCp = FcToCp(entry.StartCpOffset);
                int endCp = FcToCp(entry.EndCpOffset);
                
                int finalStart = Math.Max(0, startCp);
                int finalEnd = Math.Min(_fib.CcpText, endCp);
                
                for (int cp = finalStart; cp < finalEnd; cp++) papMap[cp] = entry.Properties;
            }
        }
        return papMap;
    }

    private PapFkp? GetPapFkp(uint pn)
    {
        if (_papFkpCache.TryGetValue(pn, out var cached)) return cached;
        var fkp = LoadPapFkp(pn);
        if (fkp != null) _papFkpCache[pn] = fkp;
        return fkp;
    }

    private PapFkp? LoadPapFkp(uint pn)
    {
        try
        {
            var offset = pn * WordConsts.FKP_PAGE_SIZE;
            if (offset + WordConsts.FKP_PAGE_SIZE > _wordDocReader.BaseStream.Length) return null;
            _wordDocReader.BaseStream.Seek(offset, SeekOrigin.Begin);
            return ParsePapFkp(_wordDocReader.ReadBytes(WordConsts.FKP_PAGE_SIZE));
        }
        catch
        {
            return null;
        }
    }

    private PapFkp ParsePapFkp(byte[] data)
    {
        var fkp = new PapFkp();
        
        // Safety check: data must be FKP_PAGE_SIZE bytes
        if (data.Length < WordConsts.FKP_PAGE_SIZE) return fkp;
        
        // Per MS-DOC spec: cpara (crun) is the LAST byte of the 512-byte FKP page
        var crun = data[WordConsts.FKP_PAGE_SIZE - 1];
        if (crun == 0 || crun > 101) return fkp;

        // FKP layout:
        //   rgfc[0..crun] : (crun+1) x 4-byte FC/CP values at offset 0
        //   rgbx[0..crun-1] : crun x 13-byte BX entries (for PAPX)
        //     Each BX: 1 byte (offset), 12 bytes (PHE descriptor)
        //   ... property data ...
        //   cpara : 1 byte at offset 511
        
        var rgfcSize = (crun + 1) * 4;
        if (rgfcSize > data.Length) return fkp;

        // Read FC/CP array (crun+1 entries, starting at offset 0)
        var fcArray = new int[crun + 1];
        for (int i = 0; i <= crun; i++)
        {
            fcArray[i] = BitConverter.ToInt32(data, i * 4);
        }

        // BX entries start right after the FC array
        // Each BX is 13 bytes for PAPX FKP (1 byte offset + 12 bytes PHE)
        var bxBase = rgfcSize;
        var bxSize = 13; // PAPX BX size

        for (int i = 0; i < crun; i++)
        {
            var bxOffset = bxBase + i * bxSize;
            if (bxOffset >= data.Length) break;
            
            var bx = data[bxOffset]; // First byte of BX is the word offset
            if (bx == 0) continue;

            var dataOffset = bx * 2;
            if (dataOffset >= WordConsts.FKP_PAGE_SIZE || dataOffset >= data.Length) continue;

            if (dataOffset + 1 > data.Length) continue;
            // cb2 is stored as a count of words (including the istd)
            var cb2 = data[dataOffset];
            var cb = cb2 * 2; // Convert word count to byte count
            if (cb == 0) continue;

            var props = new PapBase();
            if (cb >= 2 && dataOffset + 1 + cb <= data.Length)
            {
                props.Istd = BitConverter.ToUInt16(data, dataOffset + 1);
                var grpprlLength = cb - 2;
                if (grpprlLength > 0 && dataOffset + 3 + grpprlLength <= data.Length)
                {
                    var grpprl = new byte[grpprlLength];
                    Array.Copy(data, dataOffset + 3, grpprl, 0, grpprlLength);
                    // Decode paragraph and table (TAP) properties from the same GRPPRL.
                    _sprmParser.ApplyToPap(grpprl, props);
                    var tap = new TapBase();
                    _sprmParser.ApplyToTap(grpprl, tap);
                    props.Tap = tap;
                }
            }

            if (i + 1 < fcArray.Length)
            {
                fkp.Entries.Add(new PapFkpEntry
                {
                    StartCpOffset = fcArray[i],
                    EndCpOffset = fcArray[i + 1],
                    Properties = props
                });
            }
        }
        return fkp;
    }

    #endregion

    #region FC to CP conversion

    private int FcToCp(int fc)
    {
        if (_textReader == null || _textReader.Pieces.Count == 0)
        {
            // Fallback for simple documents (no Piece Table)
            // In a simple document, text starts at _fib.FcMin.
            // CPs are logical character indices starting from 0.
            if (fc < _fib.FcMin) return 0;
            
            // For simple documents, how do we know if it's 1-byte or 2-byte?
            // Windows-1252/ANSI is 1-byte. For Word 97+, text can be Unicode (2-byte) but there's a flag.
            // Simplified fallback: if it's very large, maybe it's 2-byte, but by default 
            // CP is directly FC offset for ANSI, or FC/2 for Unicode. 
            // Since ReadSimpleText already read it, let's just subtract FcMin.
            // Actually, for Word 97+, non-complex files can only be 1 byte/char if fComplex is false in some cases? 
            // The standard way: we can just assume 1 byte/char because if it's Unicode it would likely use a Piece Table.
            return (int)(fc - _fib.FcMin); 
        }

        foreach (var piece in _textReader.Pieces)
        {
            var pieceFc = (int)piece.FileOffset;
            var bytesPerChar = piece.IsUnicode ? 2 : 1;
            var pieceLengthBytes = piece.CharCount * bytesPerChar;
            
            if (fc >= pieceFc && fc < pieceFc + pieceLengthBytes)
            {
                var offsetInPiece = fc - pieceFc;
                return piece.CpStart + (offsetInPiece / bytesPerChar);
            }
        }
        
        // If not found, check if it's the exact end of the last piece
        var lastPiece = _textReader.Pieces.LastOrDefault();
        if (lastPiece != null)
        {
            var pieceFc = (int)lastPiece.FileOffset;
            var bytesPerChar = lastPiece.IsUnicode ? 2 : 1;
            var pieceLengthBytes = lastPiece.CharCount * bytesPerChar;
            
            if (fc == pieceFc + pieceLengthBytes)
            {
                return lastPiece.CpEnd;
            }
        }

        // Fallback if no matching piece is found: clamp into valid CP range
        Logger.Warning($"FKP: Unmatched FC value {fc}, clamping to document range.");
        return Math.Clamp(fc, 0, _fib.CcpText);
    }

    #endregion

    #region Convenience Methods

    public ChpBase? GetChpAtCp(int cp) => ReadChpProperties().TryGetValue(cp, out var chp) ? chp : null;
    public PapBase? GetPapAtCp(int cp) => ReadPapProperties().TryGetValue(cp, out var pap) ? pap : null;

    public RunProperties ConvertToRunProperties(ChpBase chp, StyleSheet styles)
    {
        var props = new RunProperties
        {
            FontIndex = chp.FontIndex,
            FontSize = chp.FontSize,
            FontSizeCs = chp.FontSizeCs,
            IsBold = chp.IsBold,
            IsBoldCs = chp.IsBoldCs,
            IsItalic = chp.IsItalic,
            IsItalicCs = chp.IsItalicCs,
            IsUnderline = chp.Underline != 0,
            UnderlineType = (UnderlineType)chp.Underline,
            IsStrikeThrough = chp.IsStrikeThrough,
            IsDoubleStrikeThrough = chp.IsDoubleStrikeThrough,
            IsSmallCaps = chp.IsSmallCaps,
            IsAllCaps = chp.IsAllCaps,
            IsHidden = chp.IsHidden,
            IsSuperscript = chp.IsSuperscript,
            IsSubscript = chp.IsSubscript,
            Color = chp.Color,
            CharacterSpacingAdjustment = chp.DxaOffset,
            Language = chp.LanguageId,
            // Phase 3 additions
            HighlightColor = chp.HighlightColor,
            RgbColor = chp.RgbColor,
            HasRgbColor = chp.HasRgbColor,
            IsOutline = chp.IsOutline,
            IsShadow = chp.IsShadow,
            IsEmboss = chp.IsEmboss,
            IsImprint = chp.IsImprint,
            Kerning = chp.Kerning,
            Position = chp.Position
        };
        if (chp.FontIndex >= 0 && chp.FontIndex < styles.Fonts.Count)
            props.FontName = styles.Fonts[chp.FontIndex].Name;
        return props;
    }

    public ParagraphProperties ConvertToParagraphProperties(PapBase pap, StyleSheet styles)
    {
        var styleIndex = pap.StyleId != 0 ? pap.StyleId : pap.Istd;
        return new ParagraphProperties
        {
            StyleIndex = styleIndex,
            Alignment = (ParagraphAlignment)pap.Justification,
            IndentLeft = pap.IndentLeft,
            IndentRight = pap.IndentRight,
            IndentFirstLine = pap.IndentFirstLine,
            SpaceBefore = pap.SpaceBefore,
            SpaceAfter = pap.SpaceAfter,
            LineSpacing = pap.LineSpacing,
            LineSpacingMultiple = pap.LineSpacingMultiple,
            KeepWithNext = pap.KeepWithNext,
            KeepTogether = pap.KeepTogether,
            PageBreakBefore = pap.PageBreakBefore,
            ListFormatId = pap.ListFormatId,
            ListLevel = pap.ListLevel,
            OutlineLevel = pap.OutlineLevel,
            Shading = pap.Shading
        };
    }

    #endregion
}

public class ChpFkp
{
    public List<ChpFkpEntry> Entries { get; set; } = new();
}

public class ChpFkpEntry
{
    public int StartCpOffset { get; set; }
    public int EndCpOffset { get; set; }
    public ChpBase Properties { get; set; } = new();
}

public class PapFkp
{
    public List<PapFkpEntry> Entries { get; set; } = new();
}

public class PapFkpEntry
{
    public int StartCpOffset { get; set; }
    public int EndCpOffset { get; set; }
    public PapBase Properties { get; set; } = new();
}
