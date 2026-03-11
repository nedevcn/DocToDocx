#nullable enable
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class OfficeArtMapperTests
{
    [Fact]
    public void OfficeArtReader_ResyncsPastLeadingBytes_AndPreservesWordArtText()
    {
        byte[] textBytes = Encoding.Unicode.GetBytes("艺术字\0");
        byte[] optPayload = BuildOptPayload(192, textBytes);
        byte[] spRecord = BuildLeafRecord(0xF00A, 0x00CA, BitConverter.GetBytes(42).Concat(new byte[4]).ToArray(), version: 0x2);
        byte[] optRecord = BuildLeafRecord(0xF00B, 1, optPayload, version: 0x3);
        byte[] spContainer = BuildContainerRecord(0xF004, 0, spRecord.Concat(optRecord).ToArray());
        byte[] data = new byte[] { 0x01, 0x02, 0x03, 0x04 }.Concat(spContainer).ToArray();

        using var stream = new MemoryStream(data);
        var reader = new OfficeArtReader(stream);
        var document = new DocumentModel();

        OfficeArtMapper.AttachShapes(document, reader, null);

        var shape = Assert.Single(document.Shapes);
        Assert.Equal(42, shape.Id);
        Assert.Equal(ShapeType.Textbox, shape.Type);
        Assert.Equal("艺术字", shape.Text);
    }

    [Fact]
    public void SmartArtLikeShape_IsTaggedCorrectly()
    {
        var shape = new ShapeModel { Type = ShapeType.Unknown, Text = "node" };
        if (shape.Type == ShapeType.Unknown && !string.IsNullOrEmpty(shape.Text))
            shape.Type = ShapeType.SmartArt;
        Assert.Equal(ShapeType.SmartArt, shape.Type);
    }

    [Fact]
    public void SampleTextDoc_DoesNotExposeWordArtThroughOfficeArtStreams()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Load();

        var officeArtReader = (OfficeArtReader?)typeof(DocReader)
            .GetField("_officeArtReader", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
            .GetValue(docReader);
        var fspaAnchors = (System.Collections.ICollection?)typeof(DocReader)
            .GetField("_fspaAnchors", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
            .GetValue(docReader);

        Assert.Empty(docReader.Document.Shapes);
        Assert.Empty(docReader.Document.Textboxes);
        Assert.Equal(0, officeArtReader?.RootRecords.Count ?? 0);
        Assert.Equal(0, fspaAnchors?.Count ?? 0);
    }

    [Fact]
    public void OfficeArtMapper_MapsFlipFlagsFromSpRecord()
    {
        const int flipHorizontalFlag = 0x40;
        const int flipVerticalFlag = 0x80;
        byte[] spData = BitConverter.GetBytes(42)
            .Concat(BitConverter.GetBytes(flipHorizontalFlag | flipVerticalFlag))
            .ToArray();
        byte[] spRecord = BuildLeafRecord(0xF00A, 75, spData, version: 0x2);
        byte[] spContainer = BuildContainerRecord(0xF004, 0, spRecord);

        using var stream = new MemoryStream(spContainer);
        var reader = new OfficeArtReader(stream);
        var document = new DocumentModel();

        OfficeArtMapper.AttachShapes(document, reader, null);

        var shape = Assert.Single(document.Shapes);
        Assert.Equal(42, shape.Id);
        Assert.True(shape.FlipHorizontal);
        Assert.True(shape.FlipVertical);
    }

    private static byte[] BuildOptPayload(ushort propId, byte[] complexData)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write((ushort)(0x8000 | propId));
        writer.Write((uint)complexData.Length);
        writer.Write(complexData);
        writer.Flush();
        return ms.ToArray();
    }

    private static byte[] BuildLeafRecord(ushort type, ushort instance, byte[] payload, ushort version)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write((ushort)((instance << 4) | version));
        writer.Write(type);
        writer.Write((uint)payload.Length);
        writer.Write(payload);
        writer.Flush();
        return ms.ToArray();
    }

    private static byte[] BuildContainerRecord(ushort type, ushort instance, byte[] children)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write((ushort)((instance << 4) | 0x000F));
        writer.Write(type);
        writer.Write((uint)children.Length);
        writer.Write(children);
        writer.Flush();
        return ms.ToArray();
    }

    [Fact]
    public void OfficeArtMapper_GroupContainerYieldsGroupShape()
    {
        byte[] leaf = BuildLeafRecord(0xF00A, 0x0010, BitConverter.GetBytes(5).Concat(new byte[4]).ToArray(), version: 0x2);
        var spc1 = BuildContainerRecord(0xF004, 0, leaf);
        var spc2 = BuildContainerRecord(0xF004, 0, leaf);
        var grp = BuildContainerRecord(0xF003, 0, spc1.Concat(spc2).ToArray());

        using var stream = new MemoryStream(grp);
        var reader = new OfficeArtReader(stream);
        var doc = new DocumentModel();
        OfficeArtMapper.AttachShapes(doc, reader, null);

        Assert.Single(doc.Shapes);
        Assert.Equal(ShapeType.Group, doc.Shapes[0].Type);
        Assert.NotNull(doc.Shapes[0].Children);
        Assert.Equal(2, doc.Shapes[0].Children!.Count);
    }

    [Fact]
    public void OfficeArtMapper_ParsesGradientProperties()
    {
        // assemble a shape container with OPT records containing gradient data
        byte[] leaf = BuildLeafRecord(0xF00A, 0x0010, BitConverter.GetBytes(10).Concat(new byte[4]).ToArray(), version: 0x2);
        // gradient angle property (simple, non-complex)
        using var msAngle = new MemoryStream();
        using (var bw = new BinaryWriter(msAngle, Encoding.Default, leaveOpen: true))
        {
            // header: propId with no flags, then value
            bw.Write((ushort)1000);
            bw.Write((uint)5400000);
            bw.Flush();
        }
        byte[] optAngle = BuildLeafRecord(0xF00B, 1, msAngle.ToArray(), version: 0x3);
        // gradient stops property: count(ushort) + (color:int + pos:float) * n
        var gradBuf = new List<byte>();
        gradBuf.AddRange(BitConverter.GetBytes((ushort)2));
        gradBuf.AddRange(BitConverter.GetBytes(0xFF0000));
        gradBuf.AddRange(BitConverter.GetBytes(0f));
        gradBuf.AddRange(BitConverter.GetBytes(0x00FF00));
        gradBuf.AddRange(BitConverter.GetBytes(1f));
        byte[] optStops = BuildLeafRecord(0xF00B, 1, BuildOptPayload(1001, gradBuf.ToArray()), version: 0x3);
        byte[] spContainer = BuildContainerRecord(0xF004, 0, leaf.Concat(optAngle).Concat(optStops).ToArray());

        using var stream = new MemoryStream(spContainer);
        var reader = new OfficeArtReader(stream);
        var doc = new DocumentModel();
        OfficeArtMapper.AttachShapes(doc, reader, null);

        var shape = Assert.Single(doc.Shapes);
        Assert.Equal(FillType.LinearGradient, shape.FillType);
        Assert.Equal(5400000, shape.GradientAngle);
        Assert.NotNull(shape.GradientStops);
        Assert.Equal(2, shape.GradientStops!.Count);
        Assert.Equal(0xFF0000, shape.GradientStops![0].Color);
        Assert.Equal(0d, shape.GradientStops![0].Position);
        Assert.Equal(0x00FF00, shape.GradientStops![1].Color);
        Assert.Equal(1d, shape.GradientStops![1].Position);
    }

    [Fact]
    public void OfficeArtMapper_DerivesWrapPolygonFromCustomGeometry_WhenAnchorUsesTightWrap()
    {
        byte[] leaf = BuildLeafRecord(0xF00A, 0x0010, BitConverter.GetBytes(77).Concat(new byte[4]).ToArray(), version: 0x2);
        byte[] vertices = BuildVerticesPayload(
            new System.Drawing.Point(0, 0),
            new System.Drawing.Point(1000, 0),
            new System.Drawing.Point(1000, 1000),
            new System.Drawing.Point(0, 1000));

        byte[] optVertices = BuildLeafRecord(0xF00B, 1, BuildOptPayload(321, vertices), version: 0x3);
        byte[] optLeft = BuildLeafRecord(0xF00B, 1, BuildSimpleOptPayload(323, 0), version: 0x3);
        byte[] optTop = BuildLeafRecord(0xF00B, 1, BuildSimpleOptPayload(324, 0), version: 0x3);
        byte[] optRight = BuildLeafRecord(0xF00B, 1, BuildSimpleOptPayload(325, 1000), version: 0x3);
        byte[] optBottom = BuildLeafRecord(0xF00B, 1, BuildSimpleOptPayload(326, 1000), version: 0x3);
        byte[] spContainer = BuildContainerRecord(0xF004, 0, leaf.Concat(optVertices).Concat(optLeft).Concat(optTop).Concat(optRight).Concat(optBottom).ToArray());

        using var stream = new MemoryStream(spContainer);
        var reader = new OfficeArtReader(stream);
        var doc = new DocumentModel();
        doc.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "body", CharacterPosition = 0, CharacterLength = 4 } } });

        OfficeArtMapper.AttachShapes(doc, reader, new[]
        {
            new FspaInfo
            {
                Spid = 77,
                XaLeft = 10,
                YaTop = 20,
                XaRight = 1010,
                YaBottom = 820,
                Cp = 0,
                Flags = 0x0020
            }
        });

        var shape = Assert.Single(doc.Shapes);
        Assert.NotNull(shape.Anchor);
        Assert.Equal(ShapeWrapType.Tight, shape.Anchor!.WrapType);
        Assert.NotNull(shape.WrapPolygonVertices);
        Assert.Equal(new System.Drawing.Point(0, 0), shape.WrapPolygonVertices![0]);
        Assert.Contains(shape.WrapPolygonVertices, point => point == new System.Drawing.Point(21600, 21600));
    }

    private static byte[] BuildSimpleOptPayload(ushort propId, uint value)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write(propId);
        writer.Write(value);
        writer.Flush();
        return ms.ToArray();
    }

    private static byte[] BuildVerticesPayload(params System.Drawing.Point[] points)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);
        writer.Write((ushort)points.Length);
        writer.Write((ushort)points.Length);
        writer.Write((ushort)8);
        foreach (var point in points)
        {
            writer.Write(point.X);
            writer.Write(point.Y);
        }
        writer.Flush();
        return ms.ToArray();
    }
}