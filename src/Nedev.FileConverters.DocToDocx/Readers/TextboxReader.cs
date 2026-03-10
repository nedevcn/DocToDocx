using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

public class TextboxReader
{
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly TextReader _textReader;
    private readonly FkpParser? _fkpParser;
    private readonly StyleSheet? _styles;

    public TextboxReader(BinaryReader tableReader, FibReader fib, TextReader textReader,
                         FkpParser? fkpParser = null, StyleSheet? styles = null)
    {
        _tableReader = tableReader;
        _fib = fib;
        _textReader = textReader;
        _fkpParser = fkpParser;
        _styles = styles;
    }

    public List<TextboxModel> ReadTextboxes()
    {
        var textboxes = new List<TextboxModel>();

        if (_fib.FcTxbx == 0 || _fib.LcbTxbx == 0 || _tableReader == null)
            return textboxes;

        if (!_tableReader.CanReadRange(_fib.FcTxbx, _fib.LcbTxbx))
        {
            Logger.Warning($"Skipped textboxes because PLCFTxbxBkd range 0x{_fib.FcTxbx:X}/0x{_fib.LcbTxbx:X} exceeds the Table stream.");
            return textboxes;
        }

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
        if (_fib.LcbTxbx < 12) // Minimum: PLC structure: (n+1)*4 + n*dataSize
            return textboxes;

        _tableReader.BaseStream.Seek(_fib.FcTxbx, SeekOrigin.Begin);

        var n = (int)((_fib.LcbTxbx - 4) / 12); // (n+1)*4 + n*8 = 12n + 4
        if (n <= 0) return textboxes;

        var cpArray = new int[n + 1];
        for (int i = 0; i <= n; i++) cpArray[i] = _tableReader.ReadInt32();

        // Calculate absolute CP offset for textboxes:
        // Textbox story starts after Body, Footnotes, Headers, Annotations, Endnotes
        int textboxStoryStartCp = _fib.CcpText + _fib.CcpFtn + _fib.CcpHdd + _fib.CcpAtn + _fib.CcpEdn;

        for (int i = 0; i < n; i++)
        {
            int relStart = cpArray[i];
            int relEnd = cpArray[i + 1];
            int length = relEnd - relStart;

            if (length <= 0)
            {
                Logger.Warning($"Skipped textbox entry {i} because its CP range [{relStart}, {relEnd}) is empty or reversed.");
                continue;
            }

            int absCp = textboxStoryStartCp + relStart;
            if (absCp < 0 || absCp >= _textReader.Text.Length)
            {
                Logger.Warning($"Skipped textbox entry {i} because absolute CP {absCp} falls outside the reconstructed text buffer.");
                continue;
            }

            length = Math.Min(length, _textReader.Text.Length - absCp);
            if (length <= 0)
            {
                Logger.Warning($"Skipped textbox entry {i} because no textbox story text is available at absolute CP {absCp}.");
                continue;
            }

            var textbox = new TextboxModel
            {
                Index = i + 1,
                StoryStartCharacterPosition = textboxStoryStartCp + relStart,
                StoryEndCharacterPosition = textboxStoryStartCp + relEnd,
                Width = 4320,
                Height = 2880
            };

            // Pull text from global TextReader using absolute CP
            var textboxText = _textReader.GetText(absCp, length);
            textboxText = CleanTextboxText(textboxText);

            if (!string.IsNullOrEmpty(textboxText))
            {
                var paragraphs = ParseTextboxParagraphs(textboxText, absCp);
                foreach (var para in paragraphs)
                {
                    textbox.Paragraphs.Add(para);
                }

                // Also populate Runs (flat list) from paragraphs
                foreach (var para in paragraphs)
                {
                    textbox.Runs.AddRange(para.Runs);
                }

                var firstParagraphWithProps = paragraphs.FirstOrDefault(paragraph => paragraph.Properties != null);
                if (firstParagraphWithProps?.Properties != null)
                {
                    textbox.HorizontalAlignment = firstParagraphWithProps.Properties.Alignment switch
                    {
                        ParagraphAlignment.Center => TextboxHorizontalAlignment.Center,
                        ParagraphAlignment.Right => TextboxHorizontalAlignment.Right,
                        _ => TextboxHorizontalAlignment.Left
                    };
                }
            }

            textboxes.Add(textbox);
        }

        return textboxes;
    }

    private List<ParagraphModel> ParseTextboxParagraphs(string text, int startCp)
    {
        var paragraphs = new List<ParagraphModel>();
        if (string.IsNullOrEmpty(text))
            return paragraphs;

        // Split by paragraph marks, tracking CP positions
        var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        int paraIndex = 0;
        int currentCp = startCp;

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                currentCp += line.Length + 1; // +1 for the delimiter
                continue;
            }

            var paragraph = new ParagraphModel
            {
                Index = paraIndex++,
                Type = ParagraphType.Normal
            };

            // Try to get actual CHP properties from FkpParser
            RunProperties? runProps = null;
            ParagraphProperties? paragraphProps = null;
            if (_fkpParser != null && _styles != null)
            {
                try
                {
                    var chp = _fkpParser.GetChpAtCp(currentCp);
                    if (chp != null)
                    {
                        runProps = _fkpParser.ConvertToRunProperties(chp, _styles);
                    }

                    var pap = _fkpParser.GetPapAtCp(currentCp);
                    if (pap != null)
                    {
                        paragraphProps = _fkpParser.ConvertToParagraphProperties(pap, _styles);
                    }
                }
                catch
                {
                    // Fall through to default properties
                }
            }

            paragraph.Properties = paragraphProps;

            paragraph.Runs.Add(new RunModel
            {
                Text = line.Trim(),
                CharacterPosition = currentCp,
                CharacterLength = line.Trim().Length,
                Properties = runProps ?? new RunProperties()
            });

            paragraphs.Add(paragraph);
            currentCp += line.Length + 1; // +1 for paragraph separator
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
