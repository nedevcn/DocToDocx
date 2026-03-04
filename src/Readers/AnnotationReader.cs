using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class AnnotationReader
{
    private readonly BinaryReader _annotationReader;
    private readonly FibReader _fib;

    public AnnotationReader(BinaryReader annotationReader, FibReader fib)
    {
        _annotationReader = annotationReader;
        _fib = fib;
    }

    public List<AnnotationModel> ReadAnnotations()
    {
        var annotations = new List<AnnotationModel>();

        if (_fib.FcAnot == 0 || _fib.LcbAnot == 0 || _annotationReader == null)
            return annotations;

        try
        {
            annotations = ReadAnnotationsInternal();
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read annotations", ex);
        }

        return annotations;
    }

    private List<AnnotationModel> ReadAnnotationsInternal()
    {
        var annotations = new List<AnnotationModel>();

        if (_fib.LcbAnot < 8)
            return annotations;

        _annotationReader.BaseStream.Seek(_fib.FcAnot, SeekOrigin.Begin);

        var grpprlSize = ReadGrpprl(_annotationReader, _fib.FcAnot, _fib.LcbAnot, out var pcdOffset);

        if (pcdOffset >= _fib.LcbAnot)
            return annotations;

        var pcdCount = (int)((_fib.LcbAnot - pcdOffset - 4) / 12);
        if (pcdCount <= 0)
            return annotations;

        var cps = new int[pcdCount + 1];
        for (int i = 0; i <= pcdCount; i++)
        {
            cps[i] = _annotationReader.ReadInt32();
        }

        for (int i = 0; i < pcdCount; i++)
        {
            var annotStartCp = cps[i];
            var annotEndCp = cps[i + 1];

            var annotation = new AnnotationModel
            {
                Id = $"anot_{i + 1}",
                StartCharacterPosition = annotStartCp,
                EndCharacterPosition = annotEndCp
            };

            var pcd = ReadPcd(_annotationReader);
            var annotText = ReadAnnotationText(_annotationReader, pcd, annotEndCp - annotStartCp);

            if (!string.IsNullOrEmpty(annotText))
            {
                var run = new RunModel
                {
                    Text = annotText,
                    CharacterPosition = annotStartCp,
                    CharacterLength = annotText.Length
                };
                annotation.Runs.Add(run);

                var paragraph = new ParagraphModel
                {
                    Index = 0,
                    Type = ParagraphType.Normal
                };
                paragraph.Runs.Add(run);
                annotation.Paragraphs.Add(paragraph);
            }

            annotations.Add(annotation);
        }

        ReadAnnotationAuthors(annotations);

        return annotations;
    }

    private void ReadAnnotationAuthors(List<AnnotationModel> annotations)
    {
        if (_fib.FcSttbfAtnMod == 0 || _fib.LcbSttbfAtnMod == 0)
            return;

        try
        {
            _annotationReader.BaseStream.Seek(_fib.FcSttbfAtnMod, SeekOrigin.Begin);

            var fExtend = _annotationReader.ReadUInt16();
            bool isExtended = (fExtend == 0xFFFF);

            int cData;
            if (isExtended)
            {
                cData = _annotationReader.ReadUInt16();
                _annotationReader.ReadUInt16();
            }
            else
            {
                cData = fExtend;
            }

            for (int i = 0; i < cData && i < annotations.Count; i++)
            {
                if (isExtended)
                {
                    var nameLength = _annotationReader.ReadUInt16();
                    if (nameLength > 0 && nameLength < 256)
                    {
                        var nameBytes = _annotationReader.ReadBytes(nameLength);
                        annotations[i].Author = Encoding.Unicode.GetString(nameBytes).TrimEnd('\0');
                    }
                }
                else
                {
                    var nameLength = _annotationReader.ReadByte();
                    if (nameLength > 0 && nameLength < 256)
                    {
                        var nameBytes = _annotationReader.ReadBytes(nameLength);
                        annotations[i].Author = Encoding.Default.GetString(nameBytes).TrimEnd('\0');
                    }
                }
            }
        }
        catch
        {
        }
    }

    private uint ReadGrpprl(BinaryReader reader, uint fc, uint lcb, out uint pcdOffset)
    {
        pcdOffset = 0;
        var grpprlSize = 0u;

        var startPos = reader.BaseStream.Position;
        var endPos = fc + lcb;

        while (reader.BaseStream.Position < endPos - 4)
        {
            var clxt = reader.ReadByte();

            if (clxt == 0x01)
            {
                grpprlSize = reader.ReadUInt16();
                reader.BaseStream.Seek(grpprlSize, SeekOrigin.Current);
            }
            else if (clxt == 0x02)
            {
                pcdOffset = (uint)(reader.BaseStream.Position - startPos - 1);
                reader.BaseStream.Position = startPos + pcdOffset;
                break;
            }
            else
            {
                break;
            }
        }

        return grpprlSize;
    }

    private (uint fc, bool fCompressed, ushort prm) ReadPcd(BinaryReader reader)
    {
        reader.ReadUInt16();
        var rawFc = reader.ReadUInt32();
        var fCompressed = (rawFc & 0x40000000) != 0;
        uint fc;
        if (fCompressed)
        {
            fc = (rawFc & 0x3FFFFFFF) / 2;
        }
        else
        {
            fc = rawFc & 0x3FFFFFFF;
        }
        var prm = reader.ReadUInt16();
        return (fc, fCompressed, prm);
    }

    private string ReadAnnotationText(BinaryReader reader, (uint fc, bool fCompressed, ushort prm) pcd, int length)
    {
        if (length <= 0)
            return string.Empty;

        var sb = new StringBuilder();

        try
        {
            var currentPos = reader.BaseStream.Position;
            reader.BaseStream.Seek(pcd.fc, SeekOrigin.Begin);

            if (pcd.fCompressed)
            {
                var ansiBytes = reader.ReadBytes(length);
                sb.Append(Encoding.Default.GetString(ansiBytes));
            }
            else
            {
                var unicodeBytes = reader.ReadBytes(length * 2);
                sb.Append(Encoding.Unicode.GetString(unicodeBytes));
            }

            reader.BaseStream.Position = currentPos;
        }
        catch
        {
        }

        return CleanAnnotationText(sb.ToString());
    }

    private string CleanAnnotationText(string text)
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
