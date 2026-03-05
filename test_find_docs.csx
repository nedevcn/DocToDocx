using System;
using System.IO;

string[] files = Directory.GetFiles(@"d:\Project\DocToDocx\test\docs", "*.doc");
Console.WriteLine(string.Join("\n", files));
