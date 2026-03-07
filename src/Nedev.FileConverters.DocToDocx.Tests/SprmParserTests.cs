#nullable enable
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Nedev.FileConverters.DocToDocx.Readers;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class SprmParserTests
{
    [Fact]
    public void ApplyToChp_DecodesWord97CharacterOpcodesByFullCode()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var chp = new ChpBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyChpSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x0835, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x0836, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4A43, 44);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4852, 200);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x484B, 16);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x0854, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x085C, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x085D, 1);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4863, 7);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6864, 0x12345678);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x4873, 0x0409);

        Assert.True(chp.IsBold);
        Assert.True(chp.IsItalic);
        Assert.Equal(44, chp.FontSize);
        Assert.Equal(200, chp.Scale);
        Assert.Equal(16, chp.Kerning);
        Assert.True(chp.IsImprint);
        Assert.True(chp.IsBoldCs);
        Assert.True(chp.IsItalicCs);
        Assert.Equal((ushort)7, chp.AuthorIndexDel);
        Assert.Equal(0x12345678u, chp.DateDel);
        Assert.Equal(0x0409, chp.LanguageId);
    }

    [Fact]
    public void ApplyToChp_RsidSprms_DoNotTriggerWord6ShadowOrEmbossFallbacks()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var chp = new ChpBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyChpSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6815, 0x01020304);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6816, 0x01020304);
        ApplySprm(parser, sprmType!, applyMethod!, chp, 0x6817, 0x01020304);

        Assert.False(chp.IsShadow);
        Assert.False(chp.IsEmboss);
    }

    [Fact]
    public void ApplyToChp_MetadataSprms_DoNotTriggerWord6FallbackFormatting()
    {
        using var stream = new MemoryStream();
        using var reader = new BinaryReader(stream);
        var parser = new SprmParser(reader, 0);
        var chp = new ChpBase();
        var applyMethod = typeof(SprmParser).GetMethod("ApplyChpSprm", BindingFlags.Instance | BindingFlags.NonPublic);
        var sprmType = typeof(SprmParser).GetNestedType("Sprm", BindingFlags.NonPublic);

        Assert.NotNull(applyMethod);
        Assert.NotNull(sprmType);

        var metadataSprms = new (ushort code, uint operand)[]
        {
            (0x0802, 1),
            (0x0806, 1),
            (0x080A, 1),
            (0x0811, 1),
            (0x0818, 1),
            (0x0855, 1),
            (0x0856, 1),
            (0x085A, 1),
            (0x0875, 1),
            (0x0882, 1),
            (0x4807, 1),
            (0x4867, 1),
            (0x6A09, 1),
            (0xC81A, 1),
            (0xCA57, 1),
            (0xCA62, 1),
            (0xCA89, 1)
        };

        foreach (var (code, operand) in metadataSprms)
            ApplySprm(parser, sprmType!, applyMethod!, chp, code, operand);

        Assert.False(chp.IsBold);
        Assert.False(chp.IsItalic);
        Assert.False(chp.IsOutline);
        Assert.False(chp.IsShadow);
        Assert.False(chp.IsEmboss);
        Assert.False(chp.IsImprint);
        Assert.False(chp.IsHidden);
        Assert.Equal(-1, chp.FontIndex);
        Assert.Equal(0, chp.Color);
    }

    [Fact]
    public void SampleTextDoc_ScalingRun_ComesFromFkpGrpprlWithoutCharScaleSprm()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)GetPrivateField(docReader, "_textReader")!;
        var globalChpMap = (Dictionary<int, ChpBase>)GetPrivateField(docReader, "_globalChpMap")!;
        var fullText = textReader.Text;
        var marker = "Scaling 200%";
        var markerCp = fullText.IndexOf(marker, StringComparison.Ordinal);

        Assert.True(markerCp >= 0, $"Could not find '{marker}' in sample text. Text excerpt: {TakeExcerpt(fullText, 0, 300)}");

        var scaleCp = markerCp + "Scaling ".Length;
        var piece = textReader.Pieces.FirstOrDefault(p => scaleCp >= p.CpStart && scaleCp < p.CpEnd);
        var pieceChp = textReader.GetPieceRunPropertiesAtCp(scaleCp);
        var directChp = globalChpMap.TryGetValue(scaleCp, out var direct) ? direct : null;
        var pieceModifiers = (Dictionary<ushort, byte[]>)GetPrivateField(textReader, "_piecePropertyModifiers")!;
        var pieceGrpprlHex = ResolvePieceGrpprlHex(piece, pieceModifiers);
        var fkpDetails = GetFkpEntriesForCp(docReader, scaleCp);

        var details = new StringBuilder();
        details.AppendLine($"scaleCp={scaleCp}");
        details.AppendLine($"markerCp={markerCp}");
        details.AppendLine($"piece={FormatPiece(piece)}");
        details.AppendLine($"pieceChp={FormatChp(pieceChp)}");
        details.AppendLine($"directChp={FormatChp(directChp)}");
        details.AppendLine($"pieceGrpprl={pieceGrpprlHex}");
        details.AppendLine("fkpEntries:");
        foreach (var line in fkpDetails)
            details.AppendLine(line);

        Assert.Null(piece);
        Assert.Null(pieceChp);
        Assert.NotNull(directChp);
        Assert.True(directChp!.IsBold, details.ToString());
        Assert.True(directChp.IsItalic, details.ToString());
        Assert.Equal(44, directChp.FontSize);
        Assert.Equal(7, directChp.HighlightColor);
        Assert.Equal(0x000000FFu, directChp.RgbColor);
        Assert.True(directChp.HasRgbColor);
        Assert.Equal(100, directChp.Scale);
        Assert.DoesNotContain("52 48", pieceGrpprlHex, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(fkpDetails, line => line.Contains("52 48", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(fkpDetails, line => line.Contains("43 4A", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(fkpDetails, line => line.Contains("70 68", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void SampleTextDoc_LeadingRuns_ShowWhereShadowComesFrom()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)GetPrivateField(docReader, "_textReader")!;
        var globalChpMap = (Dictionary<int, ChpBase>)GetPrivateField(docReader, "_globalChpMap")!;
        var checkpoints = new[]
        {
            0,
            textReader.Text.IndexOf("居中", StringComparison.Ordinal),
            textReader.Text.IndexOf("粗体", StringComparison.Ordinal),
            textReader.Text.IndexOf("文字Scaling 200%", StringComparison.Ordinal)
        };

        var report = new StringBuilder();
        foreach (var cp in checkpoints.Where(cp => cp >= 0).Distinct())
        {
            var directChp = globalChpMap.TryGetValue(cp, out var direct) ? direct : null;
            report.AppendLine($"cp={cp} direct={FormatChp(directChp)}");
            foreach (var line in GetFkpEntriesForCp(docReader, cp))
                report.AppendLine(line);
        }

        Assert.DoesNotContain("39 08 81", report.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SampleTextDoc_NormalStyle_DoesNotCarryShadow()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var normalStyle = docReader.Document.Styles.Styles.FirstOrDefault(s => s.StyleId == 0 || string.Equals(s.Name, "Normal", StringComparison.OrdinalIgnoreCase));

        Assert.NotNull(normalStyle);
        Assert.False(normalStyle!.RunProperties?.IsShadow ?? false);
    }

    private static void ApplySprm(SprmParser parser, Type sprmType, MethodInfo applyMethod, ChpBase chp, ushort code, uint operand)
    {
        var sprm = Activator.CreateInstance(sprmType)!;
        sprmType.GetProperty("Code")!.SetValue(sprm, code);
        sprmType.GetProperty("Operand")!.SetValue(sprm, operand);
        sprmType.GetProperty("OperandSize")!.SetValue(sprm, 0);
        applyMethod.Invoke(parser, new object[] { sprm, chp });
    }

    private static object? GetPrivateField(object instance, string fieldName)
    {
        return instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(instance);
    }

    private static string ResolvePieceGrpprlHex(Piece? piece, Dictionary<ushort, byte[]> pieceModifiers)
    {
        if (piece == null || piece.Prm == 0)
            return "<none>";

        var candidateKeys = new[]
        {
            piece.Prm,
            (ushort)(piece.Prm & 0xFFFE),
            (ushort)(piece.Prm >> 1),
            (ushort)(piece.Prm & 0x7FFF),
            (ushort)((piece.Prm & 0x7FFF) >> 1)
        };

        foreach (var key in candidateKeys)
        {
            if (pieceModifiers.TryGetValue(key, out var grpprl))
                return $"key=0x{key:X4} bytes={BitConverter.ToString(grpprl).Replace('-', ' ')}";
        }

        return $"unresolved-prm=0x{piece.Prm:X4}";
    }

    private static IReadOnlyList<string> GetFkpEntriesForCp(DocReader docReader, int cp)
    {
        var fkpParser = GetPrivateField(docReader, "_fkpParser")!;
        var cache = (IDictionary)GetPrivateField(fkpParser, "_chpFkpCache")!;
        var lines = new List<string>();

        foreach (DictionaryEntry entry in cache)
        {
            var fkp = entry.Value;
            var entriesProp = fkp!.GetType().GetProperty("Entries")!;
            var entries = (IEnumerable)entriesProp.GetValue(fkp)!;

            foreach (var item in entries)
            {
                var startCp = (int)item!.GetType().GetProperty("StartCpOffset")!.GetValue(item)!;
                var endCp = (int)item.GetType().GetProperty("EndCpOffset")!.GetValue(item)!;
                if (cp < startCp || cp >= endCp)
                    continue;

                var rawGrpprl = (byte[])item.GetType().GetProperty("RawGrpprl")!.GetValue(item)!;
                var chp = (ChpBase)item.GetType().GetProperty("Properties")!.GetValue(item)!;
                lines.Add($"pn={entry.Key} cp={startCp}..{endCp} grpprl={BitConverter.ToString(rawGrpprl).Replace('-', ' ')} chp={FormatChp(chp)}");
            }
        }

        if (lines.Count == 0)
            lines.Add("<no fkp entry for cp>");

        return lines;
    }

    private static string FormatPiece(Piece? piece)
    {
        if (piece == null)
            return "<none>";

        return $"cp={piece.CpStart}..{piece.CpEnd} prm=0x{piece.Prm:X4} offset=0x{piece.FileOffset:X8} raw=0x{piece.RawFcMasked:X8} unicode={piece.IsUnicode}";
    }

    private static string FormatChp(ChpBase? chp)
    {
        if (chp == null)
            return "<none>";

        return $"bold={chp.IsBold} italic={chp.IsItalic} underline={chp.Underline} scale={chp.Scale} kern={chp.Kerning} size={chp.FontSize} color={chp.Color} highlight={chp.HighlightColor} pos={chp.Position} lang={chp.LanguageId}";
    }

    private static string TakeExcerpt(string text, int start, int length)
    {
        if (text.Length == 0)
            return string.Empty;

        var safeStart = Math.Max(0, Math.Min(start, text.Length - 1));
        var safeLength = Math.Min(length, text.Length - safeStart);
        return text.Substring(safeStart, safeLength).Replace("\r", "\\r").Replace("\n", "\\n");
    }
}