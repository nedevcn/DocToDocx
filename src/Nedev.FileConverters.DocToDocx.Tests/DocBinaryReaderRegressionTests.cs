using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.Utils;
using Xunit;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class DocBinaryReaderRegressionTests
{
    [Fact]
    public void FibReader_UsesSpecAlignedFibRgFcLcbIndices()
    {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.Default, leaveOpen: true))
        {
            writer.Write((ushort)WordConsts.FIB_MAGIC_NUMBER);
            writer.Write((ushort)0x00D9);
            writer.Write((ushort)0);
            writer.Write((ushort)1033);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(0u);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(0u);
            writer.Write(0u);

            writer.Write((ushort)0);

            writer.Write((ushort)11);
            for (int i = 0; i < 11; i++)
            {
                writer.Write(i == 3 ? 1000 : 0);
            }

            writer.Write((ushort)76);
            for (int i = 0; i < 76; i++)
            {
                writer.Write((uint)(1000 + i * 10));
                writer.Write((uint)(2000 + i));
            }
        }

        stream.Position = 0;
        var fib = new FibReader(new BinaryReader(stream, Encoding.Default, leaveOpen: true));

        fib.Read();

        Assert.Equal((uint)1010, fib.FcStshf);
        Assert.Equal((uint)1020, fib.FcPlcffndRef);
        Assert.Equal((uint)1030, fib.FcFtn);
        Assert.Equal((uint)1110, fib.FcPlcfHdd);
        Assert.Equal((uint)1150, fib.FcSttbfFfn);
        Assert.Equal((uint)1310, fib.FcDop);
        Assert.Equal((uint)1330, fib.FcClx);
        Assert.Equal((uint)1470, fib.FcEnd);
        Assert.Equal((uint)1560, fib.FcTxbx);
        Assert.Equal((uint)1570, fib.FcPlcfFldTxbx);
        Assert.Equal((uint)1730, fib.FcPlcfLst);
        Assert.Equal((uint)1740, fib.FcPlfLfo);
    }

    [Fact]
    public void StyleReader_ReadsFfnEntriesWithoutMisalignedHeaderFields()
    {
        var fontTable = BuildFontTable(BuildFfn("Arial", "ArialAlt", family: 2, pitch: 1, trueType: true, charset: 0, weight: 700));
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcSttbfFfn), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbSttbfFfn), (uint)(fontTable.Length - 4));
        });

        using var stream = new MemoryStream(fontTable);
        using var reader = new BinaryReader(stream, Encoding.Unicode, leaveOpen: true);
        var styleReader = new StyleReader(reader, fib);

        styleReader.Read();

        var font = Assert.Single(styleReader.Styles.Fonts);
        Assert.Equal("Arial", font.Name);
        Assert.Equal("ArialAlt", font.AltName);
        Assert.Equal(2, font.Family);
        Assert.Equal(1, font.Pitch);
        Assert.Equal(0, font.Charset);
        Assert.Equal(1, font.Type);
    }

    [Fact]
    public void StyleReader_ParsesStshWhenOffsetIsZero()
    {
        var stsh = BuildStyleSheetAtOffsetZero("Title", sti: 15, sgc: 1);
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcStshf), 0u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbStshf), (uint)stsh.Length);
        });

        using var stream = new MemoryStream(stsh);
        using var reader = new BinaryReader(stream, Encoding.Unicode, leaveOpen: true);
        var styleReader = new StyleReader(reader, fib);

        styleReader.Read();

        var titleStyle = Assert.Single(styleReader.Styles.Styles.Where(style =>
            style.Type == StyleType.Paragraph &&
            string.Equals(style.Name, "Title", StringComparison.OrdinalIgnoreCase)));
        Assert.Equal(ParagraphAlignment.Center, titleStyle.ParagraphProperties?.Alignment);
        Assert.True(titleStyle.RunProperties?.IsBold);
        Assert.True(titleStyle.RunProperties?.FontSize >= 56);
    }

    [Fact]
    public void RepairFootnoteStoryLength_PersistsDerivedFootnoteCpCount()
    {
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.CcpText), 100);
            SetAutoProperty(fibReader, nameof(FibReader.CcpFtn), 0);
            SetAutoProperty(fibReader, nameof(FibReader.CcpHdd), 20);
            SetAutoProperty(fibReader, nameof(FibReader.CcpAtn), 30);
            SetAutoProperty(fibReader, nameof(FibReader.CcpEdn), 40);
            SetAutoProperty(fibReader, nameof(FibReader.CcpTxbx), 50);
            SetAutoProperty(fibReader, nameof(FibReader.CcpHdrTxbx), 60);
            SetAutoProperty(fibReader, nameof(FibReader.FcFtn), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbFtn), 16u);
        });

        using var tableStream = new MemoryStream();
        using (var writer = new BinaryWriter(tableStream, Encoding.Default, leaveOpen: true))
        {
            writer.Write(0);
            writer.Write(0);
            writer.Write(12);
            writer.Write(18);
            writer.Write((ushort)1);
            writer.Write((ushort)2);
        }
        tableStream.Position = 0;

        using var tableReader = new BinaryReader(tableStream, Encoding.Default, leaveOpen: true);
        var docReader = (DocReader)FormatterServices.GetUninitializedObject(typeof(DocReader));
        SetPrivateField(docReader, "_fibReader", fib);
        SetPrivateField(docReader, "_tableReader", tableReader);

        var method = typeof(DocReader).GetMethod("RepairFootnoteStoryLength", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(method);

        var repairedTotalCp = (int)method!.Invoke(docReader, new object[] { 300 })!;

        Assert.Equal(18, fib.CcpFtn);
        Assert.Equal(318, repairedTotalCp);
    }

    [Fact]
    public void BookmarkReader_ReadsUnicodeBookmarkNamesFromSttbfBkmk()
    {
        using var tableStream = new MemoryStream();
        using (var writer = new BinaryWriter(tableStream, Encoding.Unicode, leaveOpen: true))
        {
            writer.Write(0);
            writer.Write(5);
            writer.Write(9);
            writer.Write((ushort)1);
            writer.Write((ushort)0);

            writer.Write(9);
            writer.Write(15);

            writer.Write((ushort)0xFFFF);
            writer.Write((ushort)1);
            writer.Write((ushort)0);
            writer.Write((ushort)3);
            writer.Write(Encoding.Unicode.GetBytes("书签A"));
        }
        tableStream.Position = 0;

        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlcfBkf), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlcfBkf), 12u);
            SetAutoProperty(fibReader, nameof(FibReader.FcPlcfBkl), 16u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlcfBkl), 8u);
            SetAutoProperty(fibReader, nameof(FibReader.FcSttbfBkmk), 24u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbSttbfBkmk), 14u);
        });

        using var reader = new BinaryReader(tableStream, Encoding.Unicode, leaveOpen: true);
        using var wordReader = new BinaryReader(new MemoryStream(Array.Empty<byte>()));
        var bookmarkReader = new BookmarkReader(reader, wordReader, fib);

        bookmarkReader.Read();

        var bookmark = Assert.Single(bookmarkReader.Bookmarks);
        Assert.Equal("书签A", bookmark.Name);
        Assert.Equal(5, bookmark.StartCp);
        Assert.Equal(15, bookmark.EndCp);
    }

    [Fact]
    public void BookmarkReader_InvalidPlcfBklRange_EmitsWarningInsteadOfThrowing()
    {
        using var tableStream = new MemoryStream(new byte[32]);
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlcfBkf), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlcfBkf), 12u);
            SetAutoProperty(fibReader, nameof(FibReader.FcPlcfBkl), 24u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlcfBkl), 2u);
        });

        var diagnostics = new List<ConversionDiagnostic>();

        using var reader = new BinaryReader(tableStream, Encoding.Unicode, leaveOpen: true);
        using var wordReader = new BinaryReader(new MemoryStream(Array.Empty<byte>()));
        var bookmarkReader = new BookmarkReader(reader, wordReader, fib);

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            bookmarkReader.Read();
        }

        var bookmark = Assert.Single(bookmarkReader.Bookmarks);
        Assert.Equal(0, bookmark.StartCp);
        Assert.Equal(0, bookmark.EndCp);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("PlcfBkl range", StringComparison.Ordinal));
    }

    [Fact]
    public void AnnotationReader_InvalidPlcfandTxtRange_EmitsWarningAndReturnsEmpty()
    {
        using var tableStream = new MemoryStream(new byte[32]);
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlcfandRef), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlcfandRef), 12u);
            SetAutoProperty(fibReader, nameof(FibReader.FcPlcfandTxt), 40u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlcfandTxt), 12u);
        });

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)FormatterServices.GetUninitializedObject(typeof(Nedev.FileConverters.DocToDocx.Readers.TextReader));
        using var reader = new BinaryReader(tableStream, Encoding.Default, leaveOpen: true);
        var annotationReader = new AnnotationReader(reader, fib, textReader);
        var diagnostics = new List<ConversionDiagnostic>();

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            var annotations = annotationReader.ReadAnnotations();
            Assert.Empty(annotations);
        }

        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("PlcfandTxt range", StringComparison.Ordinal));
    }

    [Fact]
    public void TextboxReader_InvalidTableRange_EmitsWarningAndReturnsEmpty()
    {
        using var tableStream = new MemoryStream(new byte[16]);
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcTxbx), 20u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbTxbx), 16u);
        });

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)FormatterServices.GetUninitializedObject(typeof(Nedev.FileConverters.DocToDocx.Readers.TextReader));
        SetPrivateField(textReader, "_text", "sample textbox content");

        using var tableReader = new BinaryReader(tableStream, Encoding.Default, leaveOpen: true);
        var textboxReader = new TextboxReader(tableReader, fib, textReader);
        var diagnostics = new List<ConversionDiagnostic>();

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            var textboxes = textboxReader.ReadTextboxes();
            Assert.Empty(textboxes);
        }

        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("PLCFTxbxBkd range", StringComparison.Ordinal));
    }

    [Fact]
    public void TextboxReader_PreservesParagraphBoundariesWhitespaceAndCpOffsets()
    {
        const string textboxStory = "  first\r\rsecond  ";

        using var tableStream = new MemoryStream();
        using (var writer = new BinaryWriter(tableStream, Encoding.Default, leaveOpen: true))
        {
            writer.Write(0u);
            writer.Write(0);
            writer.Write(textboxStory.Length);
            writer.Write(0);
            writer.Write(0);
        }
        tableStream.Position = 0;

        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcTxbx), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbTxbx), 16u);
        });

        var textReader = (Nedev.FileConverters.DocToDocx.Readers.TextReader)FormatterServices.GetUninitializedObject(typeof(Nedev.FileConverters.DocToDocx.Readers.TextReader));
        SetPrivateField(textReader, "_text", textboxStory);

        using var tableReader = new BinaryReader(tableStream, Encoding.Default, leaveOpen: true);
        var textboxReader = new TextboxReader(tableReader, fib, textReader);

        var textbox = Assert.Single(textboxReader.ReadTextboxes());
        Assert.Equal(3, textbox.Paragraphs.Count);
        Assert.Equal("  first", Assert.Single(textbox.Paragraphs[0].Runs).Text);
        Assert.Empty(textbox.Paragraphs[1].Runs);
        Assert.Equal("second  ", Assert.Single(textbox.Paragraphs[2].Runs).Text);
        Assert.Equal(0, textbox.Paragraphs[0].Runs[0].CharacterPosition);
        Assert.Equal(9, textbox.Paragraphs[2].Runs[0].CharacterPosition);
    }

    [Fact]
    public void FieldReader_ParsesSpecialWordFieldSwitches()
    {
        var reader = new FieldReader();

        var field = reader.ParseField("DATE \\@ \"MMMM d, yyyy\" \\* MERGEFORMAT \\# \"0\"");

        Assert.NotNull(field);
        Assert.Equal("MMMM d, yyyy", field!.Switches["@"]);
        Assert.Equal("MERGEFORMAT", field.Switches["*"]);
        Assert.Equal("0", field.Switches["#"]);
    }

    [Fact]
    public void ListReader_NormalizesSequentialOverrideIds_WhenPlfLfoIsMissing()
    {
        using var tableStream = new MemoryStream(new byte[16]);
        using var tableReader = new BinaryReader(tableStream, Encoding.Unicode, leaveOpen: true);
        var listReader = new ListReader(tableReader, CreateSyntheticFibReader(_ => { }));
        SetAutoProperty(listReader, nameof(ListReader.ListFormats), new List<ListFormat>
        {
            new() { ListId = 100, Type = ListType.Simple }
        });

        var normalizeMethod = typeof(ListReader).GetMethod("NormalizeListFormatOverrides", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(normalizeMethod);

        var normalizedOverrides = (List<ListFormatOverride>)normalizeMethod!.Invoke(listReader, new object[] { new List<ListFormatOverride>() })!;

        Assert.Contains(normalizedOverrides, listOverride => listOverride.OverrideId == 1 && listOverride.ListId == 100);
        Assert.Contains(normalizedOverrides, listOverride => listOverride.OverrideId == 100 && listOverride.ListId == 100);
    }

    [Fact]
    public void ListReader_ReadPlfLfo_ParsesSequentialInstanceMappingAndLevelStartOverrides()
    {
        using var tableStream = new MemoryStream();
        using (var writer = new BinaryWriter(tableStream, Encoding.Default, leaveOpen: true))
        {
            writer.Write(new byte[4]);
            writer.Write(1);
            writer.Write(42);
            writer.Write(0u);
            writer.Write(0u);
            writer.Write((byte)1);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write(7);
            writer.Write((byte)0x10);
            writer.Write((byte)0);
            writer.Write((ushort)0);
        }

        tableStream.Position = 0;

        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlfLfo), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlfLfo), 28u);
        });

        using var tableReader = new BinaryReader(tableStream, Encoding.Default, leaveOpen: true);
        var listReader = new ListReader(tableReader, fib);
        SetAutoProperty(listReader, nameof(ListReader.ListFormats), new List<ListFormat>
        {
            new() { ListId = 42, Type = ListType.Simple }
        });

        var readPlfLfoMethod = typeof(ListReader).GetMethod("ReadPlfLfo", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(readPlfLfoMethod);

        readPlfLfoMethod!.Invoke(listReader, Array.Empty<object>());

        var overrideDefinition = Assert.Single(listReader.ListFormatOverrides.Where(listOverride => listOverride.OverrideId == 1 && listOverride.ListId == 42));
        var levelOverride = Assert.Single(overrideDefinition.Levels);
        Assert.Equal(0, levelOverride.Level);
        Assert.Equal(7, levelOverride.StartAt);
        Assert.Contains(listReader.ListFormatOverrides, listOverride => listOverride.OverrideId == 42 && listOverride.ListId == 42);
    }

    [Fact]
    public void ListReader_ReadPlfLfo_ParsesLfOlvlFormattingOverrides()
    {
        using var tableStream = new MemoryStream();
        using (var writer = new BinaryWriter(tableStream, Encoding.Unicode, leaveOpen: true))
        {
            writer.Write(new byte[4]);
            writer.Write(1);
            writer.Write(42);
            writer.Write(0u);
            writer.Write(0u);
            writer.Write((byte)1);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);

            writer.Write(7);
            writer.Write((byte)0x30);
            writer.Write((byte)0);
            writer.Write((ushort)0);

            writer.Write(9);
            writer.Write((byte)3);
            writer.Write((byte)2);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write(new byte[9]);
            writer.Write((byte)0);
            writer.Write(0);
            writer.Write(0);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((ushort)2);
            writer.Write(new[] { '\0', ')' });
        }

        tableStream.Position = 0;

        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlfLfo), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlfLfo), (uint)(tableStream.Length - 4));
        });

        using var tableReader = new BinaryReader(tableStream, Encoding.Unicode, leaveOpen: true);
        var listReader = new ListReader(tableReader, fib);
        SetAutoProperty(listReader, nameof(ListReader.ListFormats), new List<ListFormat>
        {
            new() { ListId = 42, Type = ListType.Simple }
        });

        var readPlfLfoMethod = typeof(ListReader).GetMethod("ReadPlfLfo", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(readPlfLfoMethod);

        readPlfLfoMethod!.Invoke(listReader, Array.Empty<object>());

        var overrideDefinition = Assert.Single(listReader.ListFormatOverrides.Where(listOverride => listOverride.OverrideId == 1 && listOverride.ListId == 42));
        var levelOverride = Assert.Single(overrideDefinition.Levels);
        Assert.True(levelOverride.HasFormattingOverride);
        Assert.True(levelOverride.HasStartAt);
        Assert.Equal(7, levelOverride.StartAt);
        Assert.Equal(NumberFormat.UpperLetter, levelOverride.NumberFormat);
        Assert.Equal("%1)", levelOverride.NumberText);
        Assert.Equal(2, levelOverride.Alignment);
    }

    [Fact]
    public void ListReader_ReadPlfLfo_ParsesFormattingOnlyLfOlvlOverrides()
    {
        using var tableStream = new MemoryStream();
        using (var writer = new BinaryWriter(tableStream, Encoding.Unicode, leaveOpen: true))
        {
            writer.Write(new byte[4]);
            writer.Write(1);
            writer.Write(42);
            writer.Write(0u);
            writer.Write(0u);
            writer.Write((byte)1);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);

            writer.Write(0);
            writer.Write((byte)0x20);
            writer.Write((byte)0);
            writer.Write((ushort)0);

            writer.Write(1);
            writer.Write((byte)23);
            writer.Write((byte)1);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write(new byte[9]);
            writer.Write((byte)0);
            writer.Write(0);
            writer.Write(0);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((ushort)1);
            writer.Write(new[] { '*' });
        }

        tableStream.Position = 0;

        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlfLfo), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlfLfo), (uint)(tableStream.Length - 4));
        });

        using var tableReader = new BinaryReader(tableStream, Encoding.Unicode, leaveOpen: true);
        var listReader = new ListReader(tableReader, fib);
        SetAutoProperty(listReader, nameof(ListReader.ListFormats), new List<ListFormat>
        {
            new() { ListId = 42, Type = ListType.Bullet }
        });

        var readPlfLfoMethod = typeof(ListReader).GetMethod("ReadPlfLfo", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(readPlfLfoMethod);

        readPlfLfoMethod!.Invoke(listReader, Array.Empty<object>());

        var overrideDefinition = Assert.Single(listReader.ListFormatOverrides.Where(listOverride => listOverride.OverrideId == 1 && listOverride.ListId == 42));
        var levelOverride = Assert.Single(overrideDefinition.Levels);
        Assert.True(levelOverride.HasFormattingOverride);
        Assert.False(levelOverride.HasStartAt);
        Assert.Equal(NumberFormat.Bullet, levelOverride.NumberFormat);
        Assert.Equal("*", levelOverride.NumberText);
        Assert.Equal(1, levelOverride.Alignment);
    }

    [Fact]
    public void ListReader_ReadPlfLfo_MalformedLfOlvlFormattingOverride_EmitsWarningAndKeepsStartOverride()
    {
        using var tableStream = new MemoryStream();
        using (var writer = new BinaryWriter(tableStream, Encoding.Default, leaveOpen: true))
        {
            writer.Write(new byte[4]);
            writer.Write(1);
            writer.Write(42);
            writer.Write(0u);
            writer.Write(0u);
            writer.Write((byte)1);
            writer.Write((byte)0);
            writer.Write((byte)0);
            writer.Write((byte)0);

            writer.Write(7);
            writer.Write((byte)0x30);
            writer.Write((byte)0);
            writer.Write((ushort)0);
        }

        tableStream.Position = 0;

        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlfLfo), 4u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlfLfo), (uint)(tableStream.Length - 4));
        });

        using var tableReader = new BinaryReader(tableStream, Encoding.Default, leaveOpen: true);
        var listReader = new ListReader(tableReader, fib);
        SetAutoProperty(listReader, nameof(ListReader.ListFormats), new List<ListFormat>
        {
            new() { ListId = 42, Type = ListType.Simple }
        });

        var diagnostics = new List<ConversionDiagnostic>();
        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            var readPlfLfoMethod = typeof(ListReader).GetMethod("ReadPlfLfo", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(readPlfLfoMethod);
            readPlfLfoMethod!.Invoke(listReader, Array.Empty<object>());
        }

        var overrideDefinition = Assert.Single(listReader.ListFormatOverrides.Where(listOverride => listOverride.OverrideId == 1 && listOverride.ListId == 42));
        var levelOverride = Assert.Single(overrideDefinition.Levels);
        Assert.True(levelOverride.HasStartAt);
        Assert.False(levelOverride.HasFormattingOverride);
        Assert.Equal(7, levelOverride.StartAt);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("LFOLVL formatting override", StringComparison.Ordinal));
    }

    [Fact]
    public void StyleReader_InvalidFontTableRange_FallsBackToDefaultsAndEmitsWarning()
    {
        using var stream = new MemoryStream(new byte[16]);
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcSttbfFfn), 20u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbSttbfFfn), 12u);
        });

        using var reader = new BinaryReader(stream, Encoding.Default, leaveOpen: true);
        var styleReader = new StyleReader(reader, fib);
        var diagnostics = new List<ConversionDiagnostic>();

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            styleReader.Read();
        }

        Assert.NotEmpty(styleReader.Styles.Fonts);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("Skipped font table", StringComparison.Ordinal));
    }

    [Fact]
    public void StyleReader_InvalidStshRange_EmitsWarningAndKeepsDefaultStyles()
    {
        using var stream = new MemoryStream(new byte[16]);
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcStshf), 24u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbStshf), 16u);
        });

        using var reader = new BinaryReader(stream, Encoding.Default, leaveOpen: true);
        var styleReader = new StyleReader(reader, fib);
        var diagnostics = new List<ConversionDiagnostic>();

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            styleReader.Read();
        }

        Assert.NotEmpty(styleReader.Styles.Styles);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("Skipped STSH parsing", StringComparison.Ordinal));
    }

    [Fact]
    public void SttbfHelper_TruncatedEntry_EmitsWarningAndReturnsPartialStrings()
    {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true))
        {
            writer.Write(0u);
            writer.Write((ushort)0xFFFF);
            writer.Write((ushort)2);
            writer.Write((ushort)0);
            writer.Write((ushort)1);
            writer.Write(Encoding.Unicode.GetBytes("A"));
            writer.Write((ushort)4);
            writer.Write(Encoding.Unicode.GetBytes("B"));
        }

        stream.Position = 0;
        using var reader = new BinaryReader(stream, Encoding.Unicode, leaveOpen: true);
        var diagnostics = new List<ConversionDiagnostic>();

        List<string> strings;
        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            strings = SttbfHelper.ReadSttbf(reader, 4, (uint)(stream.Length - 4));
        }

        Assert.Single(strings);
        Assert.Equal("A", strings[0]);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("declares 8 bytes of text", StringComparison.Ordinal));
    }

    [Fact]
    public void SectionReader_InvalidPlcfSedRange_EmitsWarningAndReturnsEmpty()
    {
        using var tableStream = new MemoryStream(new byte[16]);
        using var wordStream = new MemoryStream(new byte[16]);
        var fib = CreateSyntheticFibReader(fibReader =>
        {
            SetAutoProperty(fibReader, nameof(FibReader.FcPlcfSed), 20u);
            SetAutoProperty(fibReader, nameof(FibReader.LcbPlcfSed), 20u);
        });

        using var tableReader = new BinaryReader(tableStream, Encoding.Default, leaveOpen: true);
        using var wordReader = new BinaryReader(wordStream, Encoding.Default, leaveOpen: true);
        var sectionReader = new SectionReader(tableReader, wordReader, fib);
        var diagnostics = new List<ConversionDiagnostic>();

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            var sections = sectionReader.ReadSections();
            Assert.Empty(sections);
        }

        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("PlcfSed range", StringComparison.Ordinal));
    }

    [Fact]
    public void ImageReader_TruncatedImageSignature_EmitsWarningInsteadOfSilentlySkipping()
    {
        using var wordStream = new MemoryStream(new byte[16]);
        using var dataStream = new MemoryStream(new byte[]
        {
            0x42, 0x4D,
            0x10, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        });
        var fib = CreateSyntheticFibReader(_ => { });

        using var wordReader = new BinaryReader(wordStream, Encoding.Default, leaveOpen: true);
        using var dataReader = new BinaryReader(dataStream, Encoding.Default, leaveOpen: true);
        var imageReader = new ImageReader(wordReader, dataReader, fib);
        var diagnostics = new List<ConversionDiagnostic>();

        using (Logger.BeginDiagnosticCapture(diagnostics))
        {
            imageReader.ExtractImages(new DocumentModel());
        }

        Assert.Contains(diagnostics, diagnostic => diagnostic.Message.Contains("image signature", StringComparison.Ordinal) && diagnostic.Message.Contains("truncated or malformed", StringComparison.Ordinal));
    }

    private static byte[] BuildFontTable(byte[] ffn)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true);
        writer.Write(0u);
        writer.Write((ushort)1);
        writer.Write((ushort)0);
        writer.Write(ffn);
        writer.Flush();
        return stream.ToArray();
    }

    private static byte[] BuildFfn(string mainName, string altName, int family, int pitch, bool trueType, byte charset, short weight)
    {
        var names = (mainName + '\0' + altName + '\0').ToCharArray();
        var nameBytes = new byte[names.Length * 2];
        Encoding.Unicode.GetBytes(names, 0, names.Length, nameBytes, 0);
        var buffer = new byte[40 + nameBytes.Length];
        buffer[0] = (byte)(buffer.Length - 1);
        buffer[1] = (byte)((family << 4) | (trueType ? 0x04 : 0x00) | (pitch & 0x03));
        BitConverter.GetBytes(weight).CopyTo(buffer, 2);
        buffer[4] = charset;
        buffer[5] = (byte)(mainName.Length + 1);
        nameBytes.CopyTo(buffer, 40);
        return buffer;
    }

    private static byte[] BuildStyleSheetAtOffsetZero(string name, ushort sti, ushort sgc)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true);

        writer.Write((ushort)18); // cbStshi
        writer.Write((ushort)1); // cstd
        writer.Write((ushort)10); // cbSTDBaseInFile
        writer.Write(new byte[14]); // remaining STSHI bytes

        var nameBytes = Encoding.Unicode.GetBytes(name);
        var cbStd = (ushort)(10 + 2 + nameBytes.Length + 2);
        writer.Write(cbStd);
        writer.Write(sti);
        writer.Write((ushort)((0x0FFF << 4) | (sgc & 0x000F)));
        writer.Write((ushort)(0x0FFF << 4));
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write((ushort)name.Length);
        writer.Write(nameBytes);
        writer.Write((ushort)0);
        writer.Flush();

        return stream.ToArray();
    }

    private static FibReader CreateSyntheticFibReader(Action<FibReader> configure)
    {
        var fib = new FibReader(new BinaryReader(new MemoryStream(new byte[512])));
        configure(fib);
        return fib;
    }

    private static void SetAutoProperty<T>(object instance, string propertyName, T value)
    {
        var field = instance.GetType().GetField($"<{propertyName}>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        field!.SetValue(instance, value);
    }

    private static void SetPrivateField<T>(object instance, string fieldName, T value)
    {
        var field = instance.GetType().GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        field!.SetValue(instance, value);
    }
}