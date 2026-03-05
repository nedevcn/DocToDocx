using System;
using System.IO;

var path = @"d:\Project\DocToDocx\src\Writers\DocumentWriter.cs";
var text = File.ReadAllText(path);

text = text.Replace("props.AuthorIns", "props.AuthorIndexIns.ToString()");
text = text.Replace("props.AuthorDel", "props.AuthorIndexDel.ToString()");

File.WriteAllText(path, text);

var path2 = @"d:\Project\DocToDocx\src\Readers\TextReader.cs";
var text2 = File.ReadAllText(path2);

text2 = text2.Replace("/// <summary>\r\n/// Sprm (Single Property Modifier) parser.\r\n", "");

File.WriteAllText(path2, text2);
