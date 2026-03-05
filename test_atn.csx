using System;
using System.IO;

var testFile = @"d:\Project\DocToDocx\test\docs\complex.doc";
if (!File.Exists(testFile)) {
    Console.WriteLine("complex.doc not found");
    return;
}

// We just need to check if there's any file with annotations to test on.
