using System.IO;
using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class SprmParser
{
    private readonly BinaryReader _reader;
    private readonly long _endPosition;

    public SprmParser(BinaryReader reader, int length)
    {
        _reader = reader;
        _endPosition = reader.BaseStream.Position + length;
    }

    public void ApplyToChp(byte[] grpprl, ChpBase chp)
    {
        if (grpprl.Length == 0) return;
        using var ms = new MemoryStream(grpprl);
        using var reader = new BinaryReader(ms, Encoding.Default, true);
        while (ms.Position < ms.Length)
        {
            try
            {
                var sprm = ReadSprm(reader);
                ApplyChpSprm(sprm, chp);
            }
            catch (EndOfStreamException)
            {
                break;
            }
        }
    }

    public void ApplyToPap(byte[] grpprl, PapBase pap)
    {
        using var ms = new MemoryStream(grpprl);
        using var reader = new BinaryReader(ms, Encoding.Default, true);
        while (ms.Position < ms.Length)
        {
            try { var sprm = ReadSprm(reader); ApplyPapSprm(sprm, pap); }
            catch (EndOfStreamException) { break; }
        }
    }

    public void ApplyToTap(byte[] grpprl, TapBase tap)
    {
        using var ms = new MemoryStream(grpprl);
        using var reader = new BinaryReader(ms, Encoding.Default, true);
        while (ms.Position < ms.Length)
        {
            try { var sprm = ReadSprm(reader); ApplyTapSprm(sprm, tap); }
            catch (EndOfStreamException) { break; }
        }
    }

    private Sprm ReadSprm(BinaryReader reader)
    {
        var sprm = new Sprm();
        sprm.Code = reader.ReadUInt16();
        sprm.OperandSize = GetOperandSize(sprm.Code);
        switch (sprm.OperandSize)
        {
            case 1: sprm.Operand = reader.ReadByte(); break;
            case 2: sprm.Operand = reader.ReadUInt16(); break;
            case 4: sprm.Operand = reader.ReadUInt32(); break;
            case 3: sprm.Operand = (uint)(reader.ReadByte() | (reader.ReadByte() << 8) | (reader.ReadByte() << 16)); break;
            default:
                if (sprm.OperandSize == 0xFF) { var varSize = reader.ReadByte(); sprm.VariableOperand = reader.ReadBytes(varSize); }
                break;
        }
        return sprm;
    }

    private int GetOperandSize(ushort sprmCode)
    {
        var spra = (sprmCode >> 10) & 0x07;
        return spra switch { 0 => 1, 1 => 2, 2 => 4, 3 => 2, 4 => 2, 5 => 2, 6 => 4, 7 => 0xFF, _ => 1 };
    }

    private void ApplyChpSprm(Sprm sprm, ChpBase chp)
    {
        // Extract 9-bit operation code from the 16-bit Word 97 sprm Code
        var sprmCode = sprm.Code & 0x01FF;
        var sgc = (sprm.Code >> 13) & 0x07;
        
        // sgc=2 is CHP (Character Properties). Some PAP/TAP might override CHP, so relax strictness
        if (sgc != 2 && sgc != 1) return;

        switch (sprmCode)
        {
            // --- Word 97+ (16-bit) SPRM Opcodes ---
            case 0x35: chp.IsBold = sprm.Operand != 0; break; // sprmCFBold
            case 0x36: chp.IsItalic = sprm.Operand != 0; break; // sprmCFItalic
            case 0x37: chp.IsStrikeThrough = sprm.Operand != 0; break; // sprmCFStrike
            case 0x38: chp.IsOutline = sprm.Operand != 0; break; // sprmCFOutline
            case 0x39: chp.IsShadow = sprm.Operand != 0; break; // sprmCFShadow
            case 0x3A: chp.IsSmallCaps = sprm.Operand != 0; break; // sprmCFSmallCaps
            case 0x3B: chp.IsAllCaps = sprm.Operand != 0; break; // sprmCFCaps
            case 0x3C: chp.IsHidden = sprm.Operand != 0; break; // sprmCFVanish
            case 0x3E: chp.Underline = (byte)sprm.Operand; break; // sprmCKul
            case 0x43: chp.FontSize = (byte)sprm.Operand; break; // sprmCHps (half-points)
            case 0x45: chp.Position = (int)(short)sprm.Operand; break; // sprmCHpsPos
            case 0x4B: chp.Scale = (int)sprm.Operand; break; // sprmCHwcr
            case 0x4F: chp.FontIndex = (short)sprm.Operand; break; // sprmCRqftc
            case 0x5C: chp.IsBoldCs = sprm.Operand != 0; break; // sprmCFBoldBi
            case 0x5D: chp.IsItalicCs = sprm.Operand != 0; break; // sprmCFItalicBi
            case 0x61: chp.FontSizeCs = (byte)sprm.Operand; break; // sprmCHpsBi
            case 0x70: // sprmCCv (24-bit RGB)
                chp.RgbColor = sprm.Operand;
                chp.HasRgbColor = true;
                break;
            case 0x42: chp.Color = (byte)sprm.Operand; break; // sprmCIco
            case 0x0C: chp.HighlightColor = (byte)sprm.Operand; break; // sprmCHighlight
            case 0x68: chp.Language = (int)sprm.Operand; break; // sprmCRgLid0
            case 0x5E: chp.FontIndexCs = (short)sprm.Operand; break; // sprmCRqftcBi
            
            // --- Word 6 (8-bit) SPRM Opcodes (Fallbacks) ---
            case 0x02: chp.IsBold = sprm.Operand != 0; break;
            case 0x03: chp.IsItalic = sprm.Operand != 0; break;
            case 0x04: chp.IsStrikeThrough = sprm.Operand != 0; break;
            case 0x05: chp.IsUnderline = sprm.Operand != 0; break;
            case 0x06: chp.IsOutline = sprm.Operand != 0; break;
            case 0x07: chp.IsSmallCaps = sprm.Operand != 0; break;
            case 0x08: chp.IsAllCaps = sprm.Operand != 0; break;
            case 0x09: chp.IsHidden = sprm.Operand != 0; break;
            case 0x0A: chp.FontIndex = (short)sprm.Operand; break;
            case 0x0B: chp.Underline = (byte)sprm.Operand; break;
            // case 0x0C: chp.Kerning = (int)(short)sprm.Operand; break; // Conflicts with Word 97 sprmCHighlight
            case 0x0D: chp.Position = (int)(short)sprm.Operand; break;
            case 0x0E: chp.Scale = (int)sprm.Operand; break;
            case 0x11: chp.Color = (byte)sprm.Operand; break;
            case 0x12: chp.FontSize = (byte)sprm.Operand; break;
            case 0x13: chp.HighlightColor = (byte)sprm.Operand; break;
            case 0x16: chp.IsShadow = sprm.Operand != 0; break;
            case 0x17: chp.IsEmboss = sprm.Operand != 0; break;
            case 0x18: chp.IsImprint = sprm.Operand != 0; break;
            case 0x2E: chp.FontIndexCs = (short)sprm.Operand; break;
            case 0x30: chp.FontSizeCs = (byte)sprm.Operand; break;
            case 0x31: chp.IsBoldCs = sprm.Operand != 0; break;
            case 0x32: chp.IsItalicCs = sprm.Operand != 0; break;
            case 0x40: chp.IsSuperscript = sprm.Operand == 1; chp.IsSubscript = sprm.Operand == 2; break;
            case 0x41: chp.IsDoubleStrikeThrough = sprm.Operand != 0; break;
            case 0x44: chp.CharacterSpacingAdjustment = (int)sprm.Operand; break;
        }
    }

    private void ApplyPapSprm(Sprm sprm, PapBase pap)
    {
        var sprmCode = sprm.Code & 0x01FF;
        var sgc = (sprm.Code >> 13) & 0x07;
        
        // sgc=1 is PAP (Paragraph Properties)
        if (sgc != 1 && sgc != 2) return;

        switch (sprmCode)
        {
            // --- Word 97+ (16-bit) SPRM Opcodes ---
            case 0x00: pap.StyleId = (ushort)sprm.Operand; break; // sprmPIstd
            case 0x03: pap.KeepWithNext = sprm.Operand != 0; break; // sprmPFKeep
            case 0x04: pap.KeepTogether = sprm.Operand != 0; break; // sprmPFKeepFollow
            case 0x05: pap.PageBreakBefore = sprm.Operand != 0; break; // sprmPPageBreakBefore
            case 0x0B: pap.ListFormatId = (int)(short)sprm.Operand; break; // sprmPIlfo
            case 0x0A: pap.ListLevel = (byte)sprm.Operand; break; // sprmPIlvl
            case 0x0E: pap.IndentRight = (int)(short)sprm.Operand; break; // sprmPDxaRight
            case 0x0F: pap.IndentLeft = (int)(short)sprm.Operand; break; // sprmPDxaLeft
            case 0x11: pap.IndentFirstLine = (int)(short)sprm.Operand; break; // sprmPDxaLeft1
            case 0x12: pap.LineSpacing = (int)(short)(sprm.Operand & 0xFFFF); break; // sprmPDyaLine
            case 0x13: pap.SpaceBefore = (int)(short)sprm.Operand; break; // sprmPDyaBefore
            case 0x14: pap.SpaceAfter = (int)(short)sprm.Operand; break; // sprmPDyaAfter
            case 0x40: pap.OutlineLevel = (byte)sprm.Operand; break; // sprmPOutlineLvl
            case 0x61: pap.Justification = (byte)sprm.Operand; break; // sprmPJc
            
            // --- Word 6 (8-bit) SPRM Opcodes (Fallbacks) ---
            case 0x02: pap.StyleId = (ushort)sprm.Operand; break;
            case 0x15: pap.LineSpacing = (int)sprm.Operand; break;
            case 0x16: pap.SpaceBefore = (int)sprm.Operand; break;
            case 0x17: pap.SpaceAfter = (int)sprm.Operand; break;
        }
    }

    private void ApplyTapSprm(Sprm sprm, TapBase tap)
    {
        var sprmCode = sprm.Code & 0x03FF;
        var sgc = (sprm.Code >> 13) & 0x07;
        if (sgc != 3) return;
        switch (sprmCode)
        {
            // Table indent from left margin (sprmTDxaLeft)
            // Operand is a signed 16‑bit twip value.
            case 0x01:
                tap.IndentLeft = (int)(short)sprm.Operand;
                break;

            // Half of the inter‑cell gap (sprmTDxaGapHalf). The effective
            // cell spacing between two adjacent cells is typically 2 * GapHalf.
            // We also update CellSpacing when it has not been set by other
            // TAP sprms so table layout code has a single, easy source.
            case 0x02:
                tap.GapHalf = (int)(short)sprm.Operand;
                if (tap.CellSpacing == 0)
                {
                    tap.CellSpacing = tap.GapHalf * 2;
                }
                break;
            case 0x03: tap.CantSplit = sprm.Operand != 0; break; // sprmTFCantSplit
            case 0x04: tap.IsHeaderRow = sprm.Operand != 0; break; // sprmTHeader
            case 0x05: break; // sprmTTableBorders
            case 0x06: tap.Justification = (byte)sprm.Operand; break; // sprmTJc
            case 0x07: break;
            case 0x08: // sprmTDefTable - cell widths definition
                if (sprm.VariableOperand != null && sprm.VariableOperand.Length > 0)
                {
                    try
                    {
                        using var defMs = new MemoryStream(sprm.VariableOperand);
                        using var defReader = new BinaryReader(defMs);
                        var cellCount = defReader.ReadByte();
                        if (cellCount > 0 && defMs.Length >= 1 + (cellCount + 1) * 2)
                        {
                            // Read cell boundary positions (in twips)
                            var boundaries = new short[cellCount + 1];
                            for (int i = 0; i <= cellCount; i++)
                                boundaries[i] = defReader.ReadInt16();
                            
                            // Calculate cell widths from boundary differences
                            tap.CellWidths = new int[cellCount];
                            for (int i = 0; i < cellCount; i++)
                                tap.CellWidths[i] = Math.Abs(boundaries[i + 1] - boundaries[i]);
                        }
                    }
                    catch { /* ignore parse errors */ }
                }
                break;
            case 0x09: break; // sprmTSetBrc (cell borders)
            case 0x0A: break;
            case 0x0B: break;
            case 0x0C: break;
            case 0x0D: break; // sprmTShd (cell shading)
            case 0x0E: break;
            case 0x0F: break;
            case 0x10: tap.RowHeight = (int)sprm.Operand; break;
            case 0x11: tap.HeightIsExact = sprm.Operand != 0; break;
            case 0x12: break;
            case 0x13: tap.CellSpacing = (int)(short)sprm.Operand; break;
            case 0x14: tap.TableWidth = (int)(short)sprm.Operand; break;
            case 0x15: break;
            case 0x16: break;
            case 0x17: break;
            case 0x18: break;
            case 0x19: break;
            case 0x1A: break;
            case 0x1B: break;
            case 0x1C: break;
            case 0x1D: break;
            case 0x1E: break;
            case 0x1F: break;
        }
    }

    private class Sprm
    {
        public ushort Code { get; set; }
        public int OperandSize { get; set; }
        public uint Operand { get; set; }
        public byte[]? VariableOperand { get; set; }
    }
}

public class ChpBase
{
    public short FontIndex { get; set; } = -1;
    public byte FontSize { get; set; } = 24;
    public byte FontSizeCs { get; set; } = 24;
    public bool IsBold { get; set; }
    public bool IsBoldCs { get; set; }
    public bool IsItalic { get; set; }
    public bool IsItalicCs { get; set; }
    public bool IsUnderline { get; set; }
    public byte Underline { get; set; }
    public bool IsStrikeThrough { get; set; }
    public bool IsSmallCaps { get; set; }
    public bool IsAllCaps { get; set; }
    public bool IsHidden { get; set; }
    public bool IsSuperscript { get; set; }
    public bool IsSubscript { get; set; }
    public byte Color { get; set; }
    public short FontIndexCs { get; set; } = -1;
    public int CharacterSpacingAdjustment { get; set; }
    public int Language { get; set; }
    public int LanguageId { get; set; }
    public bool IsDoubleStrikeThrough { get; set; }
    public int DxaOffset { get; set; }
    // Phase 3 additions
    public bool IsOutline { get; set; }
    public int Kerning { get; set; }
    public int Position { get; set; }
    public int Scale { get; set; } = 100;
    public byte HighlightColor { get; set; }
    public bool IsShadow { get; set; }
    public bool IsEmboss { get; set; }
    public bool IsImprint { get; set; }
    public uint RgbColor { get; set; }
    public bool HasRgbColor { get; set; }
}

public class PapBase
{
    public ushort StyleId { get; set; }
    public ushort Istd { get; set; }
    public byte Justification { get; set; }
    public bool KeepWithNext { get; set; }
    public bool KeepTogether { get; set; }
    public bool PageBreakBefore { get; set; }
    public int IndentLeft { get; set; }
    public int IndentRight { get; set; }
    public int IndentFirstLine { get; set; }
    public int LineSpacing { get; set; } = 240;
    public int LineSpacingMultiple { get; set; }
    public int SpaceBefore { get; set; }
    public int SpaceAfter { get; set; }
    // Phase 3 additions
    public byte OutlineLevel { get; set; } = 9; // 9 = body text
    public int NestIndent { get; set; }
    public int ListFormatId { get; set; }
    public byte ListLevel { get; set; }
    public int ListFormatOverrideId { get; set; }
    // Associated table properties (TAP) decoded from the same GRPPRL, when present.
    public TapBase? Tap { get; set; }
}

public class TapBase
{
    public int RowHeight { get; set; }
    public bool HeightIsExact { get; set; }
    // Phase 3 additions
    /// <summary>
    /// Table justification (left/center/right) as stored in TAP.
    /// </summary>
    public byte Justification { get; set; }
    /// <summary>
    /// True when this row is marked as a header row that should repeat on each page.
    /// </summary>
    public bool IsHeaderRow { get; set; }
    /// <summary>
    /// Cell spacing in twips (total distance between cell borders).
    /// </summary>
    public int CellSpacing { get; set; }
    /// <summary>
    /// Preferred table width in twips, if specified.
    /// </summary>
    public int TableWidth { get; set; }
    /// <summary>
    /// Absolute left indent of the table from the page/column margin, in twips.
    /// </summary>
    public int IndentLeft { get; set; }
    /// <summary>
    /// Half of the inter‑cell gap (TDxaGapHalf); when present, the effective
    /// cell spacing is typically 2 * GapHalf. We keep both GapHalf and the
    /// derived CellSpacing so callers can choose the most appropriate value.
    /// </summary>
    public int GapHalf { get; set; }
    /// <summary>
    /// Per‑cell widths in twips, derived from the TAP boundary positions.
    /// </summary>
    public int[]? CellWidths { get; set; }
    /// <summary>
    /// When true, the row must not be split across pages (cantSplit).
    /// </summary>
    public bool CantSplit { get; set; }
}
