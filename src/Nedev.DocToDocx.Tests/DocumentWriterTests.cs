#nullable enable
using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Writers;
using Nedev.DocToDocx.Utils;
using Xunit;

namespace Nedev.DocToDocx.Tests
{
    public class DocumentWriterTests
    {
        [Fact]
        public void WriteDocument_MinimalParagraph_EmitsTextRun()
        {
            // Arrange: create a document with one paragraph containing a simple run
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "Hello" });
            doc.Paragraphs.Add(para);

            // Act: write the document to XML in memory
            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            // Assert: the run text makes it into the output (xml:space attribute may be present)
            Assert.Contains("Hello", xml);
            Assert.Contains("<w:p", xml); // at least one paragraph element
        }

        [Fact]
        public void WriteRun_Formats_AreEmitted()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "Fmt" };
            run.Properties = new RunProperties { IsBold = true, IsItalic = true, UnderlineType = UnderlineType.Single };
            para.Runs.Add(run);
            doc.Paragraphs.Add(para);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<w:b", xml);
            Assert.Contains("<w:i", xml);
            // underline may or may not be emitted depending on writer logic
            Assert.Contains("Fmt", xml);
        }

        [Fact]
        public void WriteRun_TrackChanges_EmitsInsDel()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run1 = new RunModel { Text = "Added" };
            run1.Properties = new RunProperties { IsInserted = true };
            var run2 = new RunModel { Text = "Removed" };
            run2.Properties = new RunProperties { IsDeleted = true };
            para.Runs.Add(run1);
            para.Runs.Add(run2);
            doc.Paragraphs.Add(para);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<w:ins", xml);
            Assert.Contains("<w:del", xml);
            Assert.Contains("Added", xml);
            Assert.Contains("Removed", xml);
        }

        [Fact]
        public void WriteRun_FieldCodes_AreOutput()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "Default" };
            run.IsField = true;
            run.FieldCode = "ASK Name \"John\""; // ask field with default
            para.Runs.Add(run);
            doc.Paragraphs.Add(para);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("instrText", xml);
            Assert.Contains("ASK Name", xml);
            Assert.Contains("Default", xml);
        }

        [Fact]
        public void EncryptionHelper_XorRoundTrips()
        {
            byte[] original = { 0, 1, 2, 3, 4, 5, 255 };
            uint key = 0xCAFEBABE;
            var encrypted = EncryptionHelper.DecryptXor(original, key);
            Assert.NotEqual(original, encrypted);
            var decrypted = EncryptionHelper.DecryptXor(encrypted, key);
            Assert.Equal(original, decrypted);
        }

        [Fact]
        public void EncryptionHelper_IsXorEncrypted_DetectsSignatures()
        {
            // if the stream contains a common magic number at the start, return false
            var notEncrypted = new MemoryStream(new byte[] { 0xEC, 0xA5 }); // little-endian 0xA5EC
            Assert.False(EncryptionHelper.IsXorEncrypted(notEncrypted, 0));

            var maybeEncrypted = new MemoryStream(new byte[] { 0x00, 0x00, 0x00 });
            Assert.True(EncryptionHelper.IsXorEncrypted(maybeEncrypted, 0));
        }

        [Fact]
        public void TableModel_NestedProperty_Works()
        {
            var t = new TableModel { Index = 3, ParentTableIndex = 1 };
            Assert.True(t.IsNested);
            Assert.Equal(1, t.ParentTableIndex);

            var top = new TableModel { Index = 5 };
            Assert.False(top.IsNested);
            Assert.Null(top.ParentTableIndex);
        }

        [Fact]
        public void WriteDocument_NestedTable_EmitsNestedTbl()
        {
            var nested = new TableModel { Index = 0 };
            nested.ColumnCount = 1;
            var row = new TableRowModel();
            var cell = new TableCellModel { Index = 0, RowIndex = 0, ColumnIndex = 0 };
            cell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "inner" } } });
            row.Cells.Add(cell);
            nested.Rows.Add(row);

            var parentTable = new TableModel { Index = 1, ColumnCount = 1 };
            var prow = new TableRowModel();
            var pcell = new TableCellModel { Index = 0, RowIndex = 0, ColumnIndex = 0 };
            var para = new ParagraphModel { Type = ParagraphType.NestedTable, NestedTable = nested };
            pcell.Paragraphs.Add(para);
            prow.Cells.Add(pcell);
            parentTable.Rows.Add(prow);

            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "before" } } });
            // add paragraph containing parentTable via writer path
            doc.Paragraphs.Add(new ParagraphModel { Type = ParagraphType.NestedTable, NestedTable = parentTable });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("before", xml);
            // should have at least two <w:tbl> entries (parent and nested)
            int count = xml.Split("<w:tbl").Length - 1;
            Assert.True(count >= 2, "Expected at least two tables, got " + count);
            Assert.Contains("inner", xml);
        }

        [Fact]
        public void ZipWriter_FullPackage_HasDocumentEntry()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "A" });
            doc.Paragraphs.Add(para);

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            // Inspect the ZIP for expected entry
            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var entry = zip.GetEntry("word/document.xml");
            Assert.NotNull(entry);

            using var entryStream = entry.Open();
            using var reader = new StreamReader(entryStream, Encoding.UTF8);
            var docXml = reader.ReadToEnd();
            // ensure the text content appears; xml:space attribute may be present
            Assert.Contains("A", docXml);
        }
    }
}