using System;
using System.IO;

var path = @"d:\Project\DocToDocx\src\Writers\DocumentWriter.cs";
var text = File.ReadAllText(path);

// 1. Add _trackChangeId field
text = text.Replace("private int _runId = 0;", "private int _runId = 0;\r\n    private int _trackChangeId = 1;");

// 2. Add WriteTrackChangeStart method
string helperMethod = @"
    private void WriteTrackChangeStart(string type, RunProperties props)
    {
        _writer.WriteStartElement(""w"", type, ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"");
        _writer.WriteAttributeString(""w"", ""id"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"", (_trackChangeId++).ToString());
        
        string author = ""Unknown Author"";
        if (type == ""ins"" && !string.IsNullOrEmpty(props.AuthorIns)) author = props.AuthorIns;
        else if (type == ""del"" && !string.IsNullOrEmpty(props.AuthorDel)) author = props.AuthorDel;
        _writer.WriteAttributeString(""w"", ""author"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"", author);
        
        uint dttm = type == ""ins"" ? props.DateIns : props.DateDel;
        if (dttm != 0)
        {
            try {
                int mint = (int)(dttm & 0x3F);
                int hr = (int)((dttm >> 6) & 0x1F);
                int dom = (int)((dttm >> 11) & 0x1F);
                int mon = (int)((dttm >> 16) & 0x0F);
                int yr = 1900 + (int)((dttm >> 20) & 0x1FF);
                var dt = new DateTime(yr, Math.Max(1, mon), Math.Max(1, dom), hr, mint, 0);
                _writer.WriteAttributeString(""w"", ""date"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"", dt.ToString(""yyyy-MM-ddTHH:mm:ssZ""));
            } catch { }
        }
    }
";

// Insert it right before WriteRun
text = text.Replace("private void WriteRun(RunModel run)", helperMethod.TrimStart() + "\r\n    private void WriteRun(RunModel run)");

// 3. Update WriteRun to wrap w:ins / w:del
string origWriteRunStart = @"    // Handle hyperlink
    if (run.IsHyperlink && !string.IsNullOrEmpty(run.HyperlinkUrl))
    {
        WriteHyperlink(run);
    }
    else
    {
        _writer.WriteStartElement(""w"", ""r"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"");";

string newWriteRunStart = origWriteRunStart.Replace(
@"        _writer.WriteStartElement(""w"", ""r"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"");",
@"        bool isIns = run.Properties?.IsInserted == true;
        bool isDel = run.Properties?.IsDeleted == true;

        if (isIns) WriteTrackChangeStart(""ins"", run.Properties!);
        else if (isDel) WriteTrackChangeStart(""del"", run.Properties!);

        _writer.WriteStartElement(""w"", ""r"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"");"
);
text = text.Replace(origWriteRunStart, newWriteRunStart);

// Close the w:ins/w:del element at the end of WriteRun
string origWriteRunEnd = @"            WriteRunText(run);
            _writer.WriteEndElement(); // w:r
        }
    }

    // Handle bookmark end";

string newWriteRunEnd = origWriteRunEnd.Replace(
@"            _writer.WriteEndElement(); // w:r
        }
    }",
@"            _writer.WriteEndElement(); // w:r
        }

        if (run.Properties?.IsInserted == true || run.Properties?.IsDeleted == true)
        {
            _writer.WriteEndElement(); // w:ins or w:del
        }
    }");
text = text.Replace(origWriteRunEnd, newWriteRunEnd);

// 4. Update WriteRunText to use w:delText for deletions
string origWriteRunTextBody = @"                if (!string.IsNullOrEmpty(part))
                {
                    _writer.WriteStartElement(""w"", ""t"", wNs);
                    if (part.StartsWith("" "") || part.EndsWith("" "") || part.Contains(""  ""))
                    {
                        _writer.WriteAttributeString(""xml"", ""space"", null, ""preserve"");
                    }
                    _writer.WriteString(part);
                    _writer.WriteEndElement();
                }";
                
string newWriteRunTextBody = origWriteRunTextBody.Replace(
    @"_writer.WriteStartElement(""w"", ""t"", wNs);",
    @"string tagName = run.Properties?.IsDeleted == true ? ""delText"" : ""t"";
                    _writer.WriteStartElement(""w"", tagName, wNs);"
);
text = text.Replace(origWriteRunTextBody, newWriteRunTextBody);

string origWriteRunTextEnd = @"        if (!string.IsNullOrEmpty(remaining))
        {
            _writer.WriteStartElement(""w"", ""t"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"");
            _writer.WriteAttributeString(""xml"", ""space"", ""http://www.w3.org/XML/1998/namespace"", ""preserve"");
            _writer.WriteString(remaining);
            _writer.WriteEndElement();
        }";

string newWriteRunTextEnd = origWriteRunTextEnd.Replace(
    @"_writer.WriteStartElement(""w"", ""t"", ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"");",
    @"string tagName = run.Properties?.IsDeleted == true ? ""delText"" : ""t"";
            _writer.WriteStartElement(""w"", tagName, ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"");"
);
text = text.Replace(origWriteRunTextEnd, newWriteRunTextEnd);

File.WriteAllText(path, text);
Console.WriteLine("Successfully modified DocumentWriter.");
