#nullable enable
using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Nedev.DocToDocx;
using Nedev.DocToDocx.Cli;

namespace Nedev.DocToDocx.Tests
{
    public class ReaderIntegrationTests
    {
        [Fact]
        public void LoadDocument_SampleDoc_HasContent()
        {
            // The test.doc file is copied to output directory by the project file.
            string path = Path.Combine(AppContext.BaseDirectory, "test.doc");
            Assert.True(File.Exists(path), "Sample document must exist in output directory");

            var doc = DocToDocxConverter.LoadDocument(path);
            Assert.NotNull(doc);
            Assert.True(doc.Paragraphs.Count > 0 || doc.Tables.Count > 0 || doc.Images.Count > 0,
                "Document should contain at least one paragraph, table or image.");
        }

        [Fact]
        public async Task Cli_ConvertsSampleDoc_CreatesDocx()
        {
            string inPath = Path.Combine(AppContext.BaseDirectory, "test.doc");
            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            // Running the CLI should not throw; it writes output file
            await Nedev.DocToDocx.Cli.Program.Main(new[] { inPath, outPath });

            Assert.True(File.Exists(outPath), "CLI should produce output document");

            // cleanup
            File.Delete(outPath);
        }

        [Fact]
        public void Convert_WithProgress_ReportsStages()
        {
            string inPath = Path.Combine(AppContext.BaseDirectory, "test.doc");
            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            var reports = new List<ConversionProgress>();
            var progress = new Progress<ConversionProgress>(p => reports.Add(p));

            DocToDocxConverter.Convert(inPath, outPath, progress);

            Assert.Contains(reports, r => r.Stage == ConversionStage.Reading);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Writing);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Complete);

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

            // copy sample doc into nested folder
            var sub = Path.Combine(tempInput, "sub");
            Directory.CreateDirectory(sub);
            File.Copy(Path.Combine(AppContext.BaseDirectory, "test.doc"), Path.Combine(sub, "a.doc"));

            await Nedev.DocToDocx.Cli.Program.Main(new[] { tempInput, tempOutput, "-r" });

            string expected = Path.Combine(tempOutput, "sub", "a.docx");
            Assert.True(File.Exists(expected), "Converted file should exist");

            Directory.Delete(tempInput, true);
            Directory.Delete(tempOutput, true);
        }

        [Fact]
        public void LoadSave_RoundTrip_PreservesBasicContent()
        {
            string inPath = Path.Combine(AppContext.BaseDirectory, "test.doc");
            string tempOut = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            var model = DocToDocxConverter.LoadDocument(inPath);
            Assert.NotNull(model);
            int paraCount = model.Paragraphs.Count;

            DocToDocxConverter.SaveDocument(model, tempOut);
            Assert.True(File.Exists(tempOut));

            var model2 = DocToDocxConverter.LoadDocument(tempOut);
            Assert.NotNull(model2);
            Assert.Equal(paraCount, model2.Paragraphs.Count);

            File.Delete(tempOut);
        }
    }
}