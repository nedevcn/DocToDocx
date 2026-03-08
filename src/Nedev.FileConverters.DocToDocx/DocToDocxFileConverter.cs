using System.IO;
using Nedev.FileConverters.Core;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.Writers;

namespace Nedev.FileConverters.DocToDocx;

/// <summary>
/// DOC to DOCX converter that implements the IFileConverter interface from Nedev.FileConverters.Core
/// This enables automatic discovery and usage through the Core infrastructure
/// </summary>
[FileConverter("doc", "docx")]
public class DocToDocxFileConverter : IFileConverter
{
    /// <summary>
    /// Converts a DOC stream to DOCX format
    /// </summary>
    /// <param name="input">Input stream containing DOC data</param>
    /// <returns>Output stream containing DOCX data</returns>
    public Stream Convert(Stream input)
    {
        if (input == null)
            throw new ArgumentNullException(nameof(input));

        var output = new MemoryStream();
        MemoryStream? bufferedInput = null;
        
        try
        {
            Stream workingInput;
            if (input.CanSeek)
            {
                input.Position = 0;
                workingInput = input;
            }
            else
            {
                bufferedInput = new MemoryStream();
                input.CopyTo(bufferedInput);
                bufferedInput.Position = 0;
                workingInput = bufferedInput;
            }
            
            using var reader = new DocReader(workingInput, password: null);
            reader.Load();
            var doc = reader.Document;
            
            using var zipWriter = new ZipWriter(output);
            var options = new Writers.DocumentWriterOptions { EnableHyperlinks = true };
            zipWriter.WriteDocument(doc, options);
            
            output.Position = 0;
            return output;
        }
        catch
        {
            output.Dispose();
            throw;
        }
        finally
        {
            bufferedInput?.Dispose();
        }
    }
}