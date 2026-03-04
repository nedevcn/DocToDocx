using System.Xml;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Writes footnotes and endnotes XML for DOCX
/// </summary>
public class FootnotesWriter
{
    private readonly XmlWriter _writer;

    public FootnotesWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    /// <summary>
    /// Writes footnotes XML
    /// </summary>
    public void WriteFootnotes(List<FootnoteModel> footnotes)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "footnotes", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // Required separator and continuation separator for better compatibility
        WriteSeparatorFootnote(-1, "separator");
        WriteSeparatorFootnote(0, "continuationSeparator");

        // Write footnotes
        foreach (var footnote in footnotes)
        {
            WriteFootnote(footnote, "footnote");
        }

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }

    private void WriteSeparatorFootnote(int id, string type)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        _writer.WriteStartElement("w", "footnote", wNs);
        _writer.WriteAttributeString("w", "type", null, type);
        _writer.WriteAttributeString("w", "id", null, id.ToString());

        _writer.WriteStartElement("w", "p", wNs);
        _writer.WriteStartElement("w", "r", wNs);
        _writer.WriteStartElement("w", "separator", wNs);
        _writer.WriteEndElement(); // w:separator
        _writer.WriteEndElement(); // w:r
        _writer.WriteEndElement(); // w:p

        _writer.WriteEndElement(); // w:footnote
    }

    /// <summary>
    /// Writes endnotes XML
    /// </summary>
    public void WriteEndnotes(List<EndnoteModel> endnotes)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "endnotes", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // Write endnotes
        foreach (var endnote in endnotes)
        {
            WriteFootnote(endnote, "endnote");
        }

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }

    private void WriteFootnote(NoteModelBase note, string type)
    {
        _writer.WriteStartElement("w", type);
        _writer.WriteAttributeString("w", "id", null, note.Index.ToString());

        // Write paragraph
        foreach (var paragraph in note.Paragraphs)
        {
            WriteParagraph(paragraph);
        }

        _writer.WriteEndElement();
    }

    private void WriteParagraph(ParagraphModel paragraph)
    {
        _writer.WriteStartElement("w", "p");

        // Write runs
        foreach (var run in paragraph.Runs)
        {
            WriteRun(run);
        }

        _writer.WriteEndElement();
    }

    private void WriteRun(RunModel run)
    {
        _writer.WriteStartElement("w", "r");

        // Write run properties if present
        if (run.Properties != null)
        {
            _writer.WriteStartElement("w", "rPr");

            if (!string.IsNullOrEmpty(run.Properties.FontName))
            {
                _writer.WriteStartElement("w", "rFonts");
                _writer.WriteAttributeString("w", "ascii", null, run.Properties.FontName);
                _writer.WriteAttributeString("w", "hAnsi", null, run.Properties.FontName);
                _writer.WriteEndElement();
            }

            if (run.Properties.FontSize > 0)
            {
                _writer.WriteStartElement("w", "sz");
                _writer.WriteAttributeString("w", "val", null, run.Properties.FontSize.ToString());
                _writer.WriteEndElement();
            }

            if (run.Properties.IsBold)
            {
                _writer.WriteStartElement("w", "b");
                _writer.WriteEndElement();
            }

            if (run.Properties.IsItalic)
            {
                _writer.WriteStartElement("w", "i");
                _writer.WriteEndElement();
            }

                var colorHex = ColorHelper.ColorToHex(run.Properties.Color);
                if (colorHex != "auto")
                {
                    _writer.WriteStartElement("w", "color");
                    _writer.WriteAttributeString("w", "val", null, colorHex);
                    _writer.WriteEndElement();
                }

            _writer.WriteEndElement();
        }

        // Write text
        _writer.WriteStartElement("w", "t");
        if (!string.IsNullOrEmpty(run.Text))
        {
            if (run.Text.StartsWith(' ') || run.Text.EndsWith(' ') || run.Text.Contains("  "))
            {
                _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
            }
            _writer.WriteString(run.Text);
        }
        _writer.WriteEndElement();

        _writer.WriteEndElement();
    }
}
