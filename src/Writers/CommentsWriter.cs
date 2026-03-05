using System.Xml;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Writes the word/comments.xml part for standard track-changes and annotations support.
/// </summary>
public class CommentsWriter
{
    private readonly XmlWriter _writer;

    public CommentsWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    public void WriteComments(DocumentModel document)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "comments", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Ensure necessary namespaces for drawing/relationships if comments have pictures
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        int commentId = 0;
        foreach (var annotation in document.Annotations)
        {
            _writer.WriteStartElement("w", "comment", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
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

            // Paragraph inside comment
            var docWriter = new DocumentWriter(_writer);
            if (annotation.Paragraphs.Count > 0)
            {
                foreach (var paragraph in annotation.Paragraphs)
                {
                    // Needs to be public or internal method, assuming WriteParagraph is available?
                    // But DocumentWriter.WriteParagraph is private. If it's private, we must duplicate logic or make it internal.
                    // For now, write a simple w:p for the first phase compatibility.
                    _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteStartElement("w", "pPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteStartElement("w", "pStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    _writer.WriteAttributeString("w", "val", null, "CommentText");
                    _writer.WriteEndElement();
                    _writer.WriteEndElement();
                    
                    foreach (var run in paragraph.Runs)
                    {
                        if (string.IsNullOrEmpty(run.Text)) continue;
                        
                        _writer.WriteStartElement("w", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        _writer.WriteStartElement("w", "rPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        _writer.WriteStartElement("w", "rStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        _writer.WriteAttributeString("w", "val", null, "CommentReference");
                        _writer.WriteEndElement();
                        _writer.WriteEndElement();
                        
                        // write text directly
                        _writer.WriteStartElement("w", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        _writer.WriteString(run.Text);
                        _writer.WriteEndElement();
                        
                        _writer.WriteEndElement(); // w:r
                    }
                    _writer.WriteEndElement(); // w:p
                }
            }
            else
            {
                // Must have at least one empty paragraph
                _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "pPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteStartElement("w", "pStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
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
    }
}
