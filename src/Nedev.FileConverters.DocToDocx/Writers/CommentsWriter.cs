using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// Writes the word/comments.xml part for standard track-changes and annotations support.
/// </summary>
public class CommentsWriter
{
    private readonly XmlWriter _writer;
    private DocumentModel? _document;
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public CommentsWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    public void WriteComments(DocumentModel document)
    {
        _document = document;
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "comments", WNs);
        
        // Ensure necessary namespaces for drawing/relationships if comments have pictures
        _writer.WriteAttributeString("xmlns", "w", null, WNs);
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        int commentId = 0;
        foreach (var annotation in document.Annotations)
        {
            _writer.WriteStartElement("w", "comment", WNs);
            _writer.WriteAttributeString("w", "id", null, commentId.ToString());
            
            if (!string.IsNullOrEmpty(annotation.Author))
            {
                _writer.WriteAttributeString("w", "author", null, annotation.Author);
            }
            if (!string.IsNullOrEmpty(annotation.Initials))
            {
                _writer.WriteAttributeString("w", "initials", null, annotation.Initials);
            }
            if (annotation.Date != default && annotation.Date > new System.DateTime(1900, 1, 1))
            {
                _writer.WriteAttributeString("w", "date", null, annotation.Date.ToString("yyyy-MM-ddTHH:mm:ssZ"));
            }

            // Write paragraphs with proper formatting
            if (annotation.Paragraphs.Count > 0)
            {
                foreach (var paragraph in annotation.Paragraphs)
                {
                    _writer.WriteStartElement("w", "p", WNs);
                    _writer.WriteStartElement("w", "pPr", WNs);
                    _writer.WriteStartElement("w", "pStyle", WNs);
                    _writer.WriteAttributeString("w", "val", null, "CommentText");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    
                    foreach (var run in paragraph.Runs)
                    {
                        if (string.IsNullOrEmpty(run.Text)) continue;
                        var safeText = DocumentWriter.SanitizeXmlString(run.Text);
                        if (string.IsNullOrEmpty(safeText)) continue;
                        
                        _writer.WriteStartElement("w", "r", WNs);
                        
                        // Write run properties if available
                        WriteCommentRunProperties(run);
                        
                        // write text
                        _writer.WriteStartElement("w", "t", WNs);
                        if (safeText.StartsWith(' ') || safeText.EndsWith(' ') || safeText.Contains("  "))
                        {
                            _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                        }
                        _writer.WriteString(safeText);
                        _writer.WriteEndElement();
                        
                        _writer.WriteEndElement(); // w:r
                    }
                    _writer.WriteEndElement(); // w:p
                }
            }
            else
            {
                // Must have at least one empty paragraph
                _writer.WriteStartElement("w", "p", WNs);
                _writer.WriteStartElement("w", "pPr", WNs);
                _writer.WriteStartElement("w", "pStyle", WNs);
                _writer.WriteAttributeString("w", "val", null, "CommentText");
                _writer.WriteEndElement();
                _writer.WriteEndElement();
                _writer.WriteEndElement(); // w:p
            }

            _writer.WriteEndElement(); // w:comment
            
            // Set mapping ID on the annotation model so DocumentWriter knows which ID to use
            annotation.Id = commentId.ToString();
            commentId++;
        }

        _writer.WriteEndElement(); // w:comments
        _writer.WriteEndDocument();
        _document = null;
    }

    private void WriteCommentRunProperties(RunModel run)
    {
        var props = run.Properties;
        if (props == null || !RunPropertiesHelper.HasRunProperties(props)) return;

        _writer.WriteStartElement("w", "rPr", WNs);
        RunPropertiesHelper.WriteRunPropertiesContent(_writer, props, includeExtended: true, _document?.Theme);
        _writer.WriteEndElement(); // w:rPr
    }
}
