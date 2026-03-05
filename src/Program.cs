using System.IO;
using Nedev.DocToDocx;

if (args.Length == 0)
{
    Console.WriteLine("Usage: Nedev.DocToDocx <input.doc> [output.docx]");
    Console.WriteLine("  If output is omitted, it defaults to input name with .docx extension.");
    return 1;
}

var inputPath = args[0];
var outputPath = args.Length > 1
    ? args[1]
    : Path.ChangeExtension(inputPath, ".docx");

if (!File.Exists(inputPath))
{
    Console.WriteLine($"Error: File not found: {inputPath}");
    return 1;
}

try
{
    DocToDocxConverter.Convert(inputPath, outputPath);
    return 0;
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    return 1;
}
