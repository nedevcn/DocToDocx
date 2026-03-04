using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Header/Footer reader - parses PlcfHdd structure from the Table stream.
///
/// Based on MS-DOC specification §2.7.
///
/// PlcfHdd structure contains:
///   - Array of CP (character position) boundaries
///   - Array of Hdd (header/footer descriptor) entries
///
/// Each section can have up to 6 header/footer types:
///   - First page header/footer
///   - Odd page header/footer (default)
///   - Even page header/footer
/// </summary>
public class HeaderFooterReader
{
    private readonly BinaryReader _tableReader;
    private readonly BinaryReader _wordDocReader;
    private readonly FibReader _fib;
    private readonly TextReader _textReader;

    public List<HeaderFooterModel> Headers { get; private set; } = new();
    public List<HeaderFooterModel> Footers { get; private set; } = new();

    public HeaderFooterReader(
        BinaryReader tableReader,
        BinaryReader wordDocReader,
        FibReader fib,
        TextReader textReader)
    {
        _tableReader = tableReader;
        _wordDocReader = wordDocReader;
        _fib = fib;
        _textReader = textReader;
    }

    /// <summary>
    /// Reads header/footer information from the document.
    /// </summary>
    public void Read(DocumentModel document)
    {
        if (_fib.FcPlcfHdd == 0 || _fib.LcbPlcfHdd == 0)
        {
            // No header/footer data
            return;
        }

        try
        {
            ReadPlcfHdd(document);
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read headers/footers", ex);
        }
    }

    /// <summary>
    /// Reads PlcfHdd structure.
    ///
    /// Structure:
    ///   - Array of CPs (character positions) - (n+1) entries
    ///   - Array of Hdd entries - n entries
    ///
    /// Each Hdd entry contains 6 words (12 bytes) indicating header/footer story lengths.
    /// </summary>
    private void ReadPlcfHdd(DocumentModel document)
    {
        _tableReader.BaseStream.Seek(_fib.FcPlcfHdd, SeekOrigin.Begin);

        // Calculate number of entries
        // PlcfHdd structure: (n+1) CPs + n * 12 bytes of Hdd data
        // Total size = (n+1) * 4 + n * 12 = 4n + 4 + 12n = 16n + 4
        // So n = (size - 4) / 16

        var dataSize = (int)_fib.LcbPlcfHdd;
        if (dataSize < 8) return;

        var entryCount = (dataSize - 4) / 16;
        if (entryCount <= 0) return;

        // Read CP array
        var cpArray = new int[entryCount + 1];
        for (int i = 0; i <= entryCount; i++)
        {
            cpArray[i] = _tableReader.ReadInt32();
        }

        // Read Hdd entries
        for (int i = 0; i < entryCount; i++)
        {
            // Each Hdd entry has 6 words:
            //   cchHdd (6 words) - character counts for each header/footer type
            var hddData = new short[6];
            for (int j = 0; j < 6; j++)
            {
                hddData[j] = _tableReader.ReadInt16();
            }

            // Process this section's headers/footers
            ProcessHddEntry(i, cpArray, hddData, document);
        }
    }

    /// <summary>
    /// Processes a single Hdd entry to extract header/footer content.
    ///
    /// Hdd entry format (6 words):
    ///   Word 0: cchHeaderFirst - first page header length
    ///   Word 1: cchFooterFirst - first page footer length
    ///   Word 2: cchHeaderOdd - odd page header length
    ///   Word 3: cchFooterOdd - odd page footer length
    ///   Word 4: cchHeaderEven - even page header length
    ///   Word 5: cchFooterEven - even page footer length
    /// </summary>
    private void ProcessHddEntry(int sectionIndex, int[] cpArray, short[] hddData, DocumentModel document)
    {
        // Map of header/footer types
        var types = new[]
        {
            (Type: HeaderFooterType.HeaderFirst, Length: hddData[0]),
            (Type: HeaderFooterType.FooterFirst, Length: hddData[1]),
            (Type: HeaderFooterType.HeaderOdd, Length: hddData[2]),
            (Type: HeaderFooterType.FooterOdd, Length: hddData[3]),
            (Type: HeaderFooterType.HeaderEven, Length: hddData[4]),
            (Type: HeaderFooterType.FooterEven, Length: hddData[5])
        };

        // Calculate starting CP for this section's header/footer stories
        // The header/footer text is stored in the WordDocument stream
        // at specific CP positions

        int currentCp = cpArray[sectionIndex];

        foreach (var (type, length) in types)
        {
            if (length <= 0) continue;

            try
            {
                // Extract header/footer text
                var text = ExtractHeaderFooterText(currentCp, length);

                var model = new HeaderFooterModel
                {
                    Type = type,
                    SectionIndex = sectionIndex,
                    Text = text,
                    CharacterPosition = currentCp,
                    CharacterLength = length
                };

                if (type == HeaderFooterType.HeaderFirst ||
                    type == HeaderFooterType.HeaderOdd ||
                    type == HeaderFooterType.HeaderEven)
                {
                    Headers.Add(model);
                }
                else
                {
                    Footers.Add(model);
                }

                currentCp += length;
            }
            catch (Exception ex)
            {
                Logger.Warning("Failed to extract header/footer", ex);
            }
        }
    }

    /// <summary>
    /// Extracts header/footer text from the global text stream.
    /// Header/footer text is stored after the main document and footnote text.
    /// </summary>
    private string ExtractHeaderFooterText(int cp, int length)
    {
        if (length <= 0)
            return string.Empty;

        // The header/footer text starts at CP position (CcpText + CcpFtn) in the global stream
        int headerStoryStartCp = _fib.CcpText + _fib.CcpFtn;
        int absoluteCp = headerStoryStartCp + cp;

        string rawText = _textReader.GetText(absoluteCp, length);
        return CleanHeaderFooterText(rawText);
    }

    /// <summary>
    /// Cleans header/footer text by removing control characters.
    /// </summary>
    private string CleanHeaderFooterText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            // Skip invalid XML characters (0x00-0x1F except tab, newline, carriage return)
            if (ch < 0x09 || (ch > 0x0D && ch < 0x20))
            {
                continue;
            }
            // Skip special Word characters
            switch (ch)
            {
                case '\x01':  // Field begin mark
                case '\x13': // Field separator
                case '\x14': // Field end
                case '\x15': // Object anchor
                    continue;
                case '\x0B':
                    sb.Append('\n');
                    break;
                case '\x07':
                    sb.Append('\t');
                    break;
                case '\x1E':
                    sb.Append('-');
                    break;
                case '\x1F':
                    break;
                default:
                    sb.Append(ch);
                    break;
            }
        }

        return sb.ToString().Trim();
    }

    /// <summary>
    /// Gets headers for a specific section.
    /// </summary>
    public List<HeaderFooterModel> GetHeadersForSection(int sectionIndex)
    {
        return Headers.Where(h => h.SectionIndex == sectionIndex).ToList();
    }

    /// <summary>
    /// Gets footers for a specific section.
    /// </summary>
    public List<HeaderFooterModel> GetFootersForSection(int sectionIndex)
    {
        return Footers.Where(f => f.SectionIndex == sectionIndex).ToList();
    }
}


