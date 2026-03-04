using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

public class ListReader
{
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;

    public List<NumberingDefinition> NumberingDefinitions { get; private set; } = new();
    public List<ListFormat> ListFormats { get; private set; } = new();

    public ListReader(BinaryReader tableReader, FibReader fib)
    {
        _tableReader = tableReader;
        _fib = fib;
    }

    public void Read()
    {
        if (_fib.FcPlcfLst == 0 || _fib.LcbPlcfLst == 0)
        {
            return;
        }

        try
        {
            ReadPlcfLst();
            ReadPlfLfo();
            BuildNumberingDefinitions();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Failed to read list formats: {ex.Message}");
        }
    }

    private void BuildNumberingDefinitions()
    {
        NumberingDefinitions = new List<NumberingDefinition>();

        foreach (var listFormat in ListFormats)
        {
            var numDef = new NumberingDefinition
            {
                Id = listFormat.ListId
            };

            for (int i = 0; i < Math.Min(listFormat.Levels.Count, 9); i++)
            {
                var lvl = listFormat.Levels[i];
                var numLevel = new NumberingLevel
                {
                    Level = lvl.Level,
                    NumberFormat = lvl.NumberFormat,
                    Text = lvl.NumberText ?? "%" + (i + 1),
                    Start = lvl.StartAt,
                    ParagraphProperties = lvl.ParagraphProperties,
                    RunProperties = lvl.RunProperties
                };
                numDef.Levels.Add(numLevel);
            }

            if (numDef.Levels.Count == 0)
            {
                for (int i = 0; i < 9; i++)
                {
                    numDef.Levels.Add(new NumberingLevel
                    {
                        Level = i,
                        NumberFormat = NumberFormat.Decimal,
                        Text = "%" + (i + 1),
                        Start = 1
                    });
                }
            }

            NumberingDefinitions.Add(numDef);
        }
    }

    private void ReadPlcfLst()
    {
        _tableReader.BaseStream.Seek(_fib.FcPlcfLst, SeekOrigin.Begin);
        var endPos = _fib.FcPlcfLst + _fib.LcbPlcfLst;

        if (endPos - _tableReader.BaseStream.Position < 4)
            return;

        var lstfCount = _tableReader.ReadInt32();
        if (lstfCount <= 0 || lstfCount > 1000)
            lstfCount = 64;

        var lists = new List<ListFormat>();

        for (int i = 0; i < lstfCount && _tableReader.BaseStream.Position + 20 <= endPos; i++)
        {
            try
            {
                var listFormat = ReadLstf(endPos);
                if (listFormat != null)
                {
                    lists.Add(listFormat);
                }
            }
            catch
            {
                break;
            }
        }

        ListFormats = lists;
    }

    private ListFormat? ReadLstf(long endPos)
    {
        var startPos = _tableReader.BaseStream.Position;

        if (endPos - startPos < 28)
            return null;

        var lsid = _tableReader.ReadInt32();
        if (lsid == 0)
            return null;

        var tplc = _tableReader.ReadUInt32();

        var rgistd = new ushort[9];
        for (int i = 0; i < 9; i++)
        {
            if (_tableReader.BaseStream.Position + 2 <= endPos)
            {
                rgistd[i] = _tableReader.ReadUInt16();
            }
        }

        var flags = _tableReader.ReadUInt16();

        var styleIndex = (ushort)((flags >> 4) & 0x0F);
        var listType = (ListType)(flags & 0x03);

        var listFormat = new ListFormat
        {
            ListId = lsid,
            Type = listType
        };

        for (int lvl = 0; lvl < 9; lvl++)
        {
            var listLevel = new ListLevel
            {
                Level = lvl,
                NumberFormat = lvl == 0 && listType == ListType.Bullet ? NumberFormat.Bullet : NumberFormat.Decimal,
                StartAt = 1,
                Indent = 720 * (lvl + 1),
                NumberText = lvl == 0 && listType == ListType.Bullet ? "·" : "%" + (lvl + 1)
            };
            listFormat.Levels.Add(listLevel);
        }

        return listFormat;
    }

    private void ReadPlfLfo()
    {
        if (_fib.FcPlfLfo == 0 || _fib.LcbPlfLfo == 0)
            return;

        _tableReader.BaseStream.Seek(_fib.FcPlfLfo, SeekOrigin.Begin);
        var endPos = _fib.FcPlfLfo + _fib.LcbPlfLfo;

        if (endPos - _tableReader.BaseStream.Position < 4)
            return;

        var lfoCount = _tableReader.ReadInt32();
        if (lfoCount <= 0 || lfoCount > 1000)
            return;

        for (int i = 0; i < lfoCount && _tableReader.BaseStream.Position + 20 <= endPos; i++)
        {
            try
            {
                var lsbfr = _tableReader.ReadInt32();
                var reserved = _tableReader.ReadInt32();

                var flags = _tableReader.ReadUInt16();
                _tableReader.ReadUInt16();

                var grpLfo = new byte[12];
                for (int j = 0; j < 12 && _tableReader.BaseStream.Position < endPos; j++)
                {
                    grpLfo[j] = _tableReader.ReadByte();
                }

                if (i < ListFormats.Count)
                {
                    ApplyLfoOverrides(ListFormats[i], grpLfo);
                }
            }
            catch
            {
                break;
            }
        }
    }

    private void ApplyLfoOverrides(ListFormat listFormat, byte[] grpLfo)
    {
        if (grpLfo.Length < 12)
            return;

        for (int lvl = 0; lvl < Math.Min(listFormat.Levels.Count, 9); lvl++)
        {
            var offset = lvl * 2;
            if (offset + 1 < grpLfo.Length)
            {
                var startAt = grpLfo[offset] | (grpLfo[offset + 1] << 8);
                if (startAt > 0)
                {
                    listFormat.Levels[lvl].StartAt = startAt;
                }
            }
        }
    }

    public ListFormat? GetListFormat(int listId)
    {
        return ListFormats.FirstOrDefault(l => l.ListId == listId);
    }

    public static bool IsListParagraph(ParagraphModel paragraph)
    {
        if (paragraph.Properties == null) return false;
        return paragraph.ListFormatId > 0 || paragraph.ListLevel > 0;
    }

    public int GetListLevel(ParagraphModel paragraph)
    {
        return paragraph.ListLevel;
    }
}
