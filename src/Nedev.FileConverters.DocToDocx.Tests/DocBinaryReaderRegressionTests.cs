using System;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.Utils;
using Xunit;

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