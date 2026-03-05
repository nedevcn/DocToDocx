#nullable enable
using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Nedev.DocToDocx;
using Nedev.DocToDocx.Cli;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Tests
{
    public class ReaderIntegrationTests
    {
        [Fact]
        public void CreateAndLoadDocument_HasContent()
        {
            // LoadDocument only supports .doc files; attempting to load a .docx should throw.
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var original = new DocumentModel();
            original.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Hello" } } });
            DocToDocxConverter.SaveDocument(original, path);

            try
            {
                Assert.Throws<InvalidDataException>(() => DocToDocxConverter.LoadDocument(path));
            }
            finally
            {
                // ensure the temporary file is removed even if the assertion fails or a handle was leaked
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [Fact]
        public async Task Cli_CopiesDocxInput_WhenPassedDocx()
        {
            string tempInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "X" } } });
            DocToDocxConverter.SaveDocument(doc, tempInput);

            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            await Nedev.DocToDocx.Cli.Program.Main(new[] { tempInput, outPath });
            Assert.True(File.Exists(outPath));

            // verify copy semantics (size)
            Assert.Equal(new FileInfo(tempInput).Length, new FileInfo(outPath).Length);
            File.Delete(tempInput);
            File.Delete(outPath);
        }

        [Fact]
        public void Convert_WithProgress_ReportsStages()
        {
            // create a simple docx file instead of relying on an external .doc sample
            string inPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var model = new DocumentModel();
            model.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "progress" } } });
            DocToDocxConverter.SaveDocument(model, inPath);

            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            var reports = new List<ConversionProgress>();
            var progress = new Progress<ConversionProgress>(p => reports.Add(p));

            // since we pass a .docx file, the converter will copy it; progress stages should still fire.
            DocToDocxConverter.Convert(inPath, outPath, progress);

            Assert.Contains(reports, r => r.Stage == ConversionStage.Reading);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Writing);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Complete);

            File.Delete(outPath);
            File.Delete(inPath);
        }

        [Fact]
        public void Convert_WithoutProgress_CopiesDocx()
        {
            string inPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "copy" } } });
            DocToDocxConverter.SaveDocument(doc, inPath);

            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            DocToDocxConverter.Convert(inPath, outPath);

            Assert.True(File.Exists(outPath));
            Assert.Equal(new FileInfo(inPath).Length, new FileInfo(outPath).Length);

            File.Delete(inPath);
            File.Delete(outPath);
        }

        [Fact]
        public async Task Cli_VersionFlag_PrintsVersion()
        {
            using var sw = new StringWriter();
            Console.SetOut(sw);
            await Nedev.DocToDocx.Cli.Program.Main(new[] { "--version" });
            string output = sw.ToString();
            Assert.Contains("Version", output);
        }

        [Fact]
        public async Task Cli_DirectoryConversion_WritesFiles()
        {
            string tempInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            string tempOutput = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempInput);
            Directory.CreateDirectory(tempOutput);

            // create a nested docx file
            var sub = Path.Combine(tempInput, "sub");
            Directory.CreateDirectory(sub);
            string aPath = Path.Combine(sub, "a.docx");
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Z" } } });
            DocToDocxConverter.SaveDocument(doc, aPath);

            await Nedev.DocToDocx.Cli.Program.Main(new[] { tempInput, tempOutput, "-r" });

            string expected = Path.Combine(tempOutput, "sub", "a.docx");
            Assert.True(File.Exists(expected), "Converted file should exist");

            Directory.Delete(tempInput, true);
            Directory.Delete(tempOutput, true);
        }

        [Fact]
        public void SaveDocument_GeneratesValidDocx()
        {
            var model = new DocumentModel();
            model.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Hello" } } });
            string tempOut = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            DocToDocxConverter.SaveDocument(model, tempOut);
            Assert.True(File.Exists(tempOut));

            using (var zip = new System.IO.Compression.ZipArchive(File.OpenRead(tempOut), System.IO.Compression.ZipArchiveMode.Read))
            {
                var entry = zip.GetEntry("word/document.xml");
                Assert.NotNull(entry);
                using var reader = new StreamReader(entry.Open());
                var xml = reader.ReadToEnd();
                Assert.Contains("Hello", xml);
            }

            File.Delete(tempOut);
        }
    }
}