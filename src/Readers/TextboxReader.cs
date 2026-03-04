using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class TextboxReader
{
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly TextReader _textReader;

    public TextboxReader(BinaryReader tableReader, FibReader fib, TextReader textReader)
    {
        _tableReader = tableReader;
        _fib = fib;
        _textReader = textReader;
    }

    public List<TextboxModel> ReadTextboxes()
    {
        var textboxes = new List<TextboxModel>();

        if (_fib.FcTxbx == 0 || _fib.LcbTxbx == 0 || _tableReader == null)
            return textboxes;

        try
        {
            textboxes = ReadTextboxesInternal();
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read textboxes", ex);
        }

        return textboxes;
    }

    private List<TextboxModel> ReadTextboxesInternal()
    {
        var textboxes = new List<TextboxModel>();

        // PLCFTxbxBkd (fcTxbx) contains boundaries in the textbox story
        // Each entry is 8 bytes (FTXBX)
        if (_fib.LcbTxbx < 12) // Minimum: 2 CPs (8 bytes) + 1 FTXBX (8 bytes) = 16 bytes? 
                                // Actually PLC structure: (n+1)*4 + n*dataSize
            return textboxes;

        _tableReader.BaseStream.Seek(_fib.FcTxbx, SeekOrigin.Begin);

        var n = (int)((_fib.LcbTxbx - 4) / 12); // (n+1)*4 + n*8 = 12n + 4
        if (n <= 0) return textboxes;

        var cpArray = new int[n + 1];
        for (int i = 0; i <= n; i++) cpArray[i] = _tableReader.ReadInt32();

        // Skip FTXBX descriptors for now (or read if needed)
        // _tableReader.BaseStream.Seek(n * 8, SeekOrigin.Current);

        // Calculate absolute CP offset for textboxes:
        // Textbox story starts after Body, Footnotes, Headers, Annotations, Endnotes
        int textboxStoryStartCp = _fib.CcpText + _fib.CcpFtn + _fib.CcpHdd + _fib.CcpAtn + _fib.CcpEdn;

        for (int i = 0; i < n; i++)
        {
            int relStart = cpArray[i];
            int relEnd = cpArray[i + 1];
            int length = relEnd - relStart;

            if (length <= 0) continue;

            var textbox = new TextboxModel
            {
                Index = i + 1,
                Width = 4320,
                Height = 2880
            };

            // Pull text from global TextReader using absolute CP
            var textboxText = _textReader.GetText(textboxStoryStartCp + relStart, length);

            if (!string.IsNullOrEmpty(textboxText))
            {
                var runs = ParseTextboxRuns(textboxText, textboxStoryStartCp + relStart);
                textbox.Runs.AddRange(runs);

                var paragraphs = ParseTextboxParagraphs(textboxText);
                foreach (var para in paragraphs)
                {
                    textbox.Paragraphs.Add(para);
                }
            }

            textboxes.Add(textbox);
        }

        return textboxes;
    }

    private List<RunModel> ParseTextboxRuns(string text, int startCp)
    {
        var runs = new List<RunModel>();
        if (string.IsNullOrEmpty(text))
            return runs;

        var paragraphs = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        int cp = startCp;

        foreach (var para in paragraphs)
        {
            if (!string.IsNullOrWhiteSpace(para))
            {
                runs.Add(new RunModel
                {
                    Text = para.Trim(),
                    CharacterPosition = cp,
                    CharacterLength = para.Trim().Length,
                    Properties = new RunProperties { FontSize = 24 }
                });
                cp += para.Length;
            }
        }

        if (runs.Count == 0 && !string.IsNullOrWhiteSpace(text))
        {
            runs.Add(new RunModel
            {
                Text = text.Trim(),
                CharacterPosition = startCp,
                CharacterLength = text.Length,
                Properties = new RunProperties { FontSize = 24 }
            });
        }

        return runs;
    }

    private List<ParagraphModel> ParseTextboxParagraphs(string text)
    {
        var paragraphs = new List<ParagraphModel>();
        if (string.IsNullOrEmpty(text))
            return paragraphs;

        var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        int paraIndex = 0;

        foreach (var line in lines)
        {
            if (!string.IsNullOrWhiteSpace(line))
            {
                var paragraph = new ParagraphModel
                {
                    Index = paraIndex++,
                    Type = ParagraphType.Normal
                };

                paragraph.Runs.Add(new RunModel
                {
                    Text = line.Trim(),
                    Properties = new RunProperties { FontSize = 24 }
                });

                paragraphs.Add(paragraph);
            }
        }

        return paragraphs;
    }

    private string CleanTextboxText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '\x01':
                case '\x13':
                case '\x14':
                case '\x15':
                    break;
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
}
