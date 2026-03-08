using System.Buffers.Binary;
using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Reads style definitions from Word 97-2003 documents.
/// Parses the STSH (Style Sheet) structure from the Table stream.
/// 
/// Phase 1: Reads font table (SttbfFfn) from Table stream.
///          Uses simplified/default style definitions.
/// Phase 2 will implement full STD parsing from STSH.
/// </summary>
public class StyleReader
{
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    public StyleSheet Styles { get; private set; } = new();

    public StyleReader(BinaryReader tableReader, FibReader fib)
    {
        _tableReader = tableReader;
        _fib = fib;
    }

    /// <summary>
    /// Reads the style sheet and font table.
    /// </summary>
    public void Read()
    {
        ReadFontTable();
        ReadStyleDefinitions();
    }

    /// <summary>
    /// Reads the font table (SttbfFfn) from the Table stream.
    /// The font table is located at fcSttbfFfn in the Table stream.
    /// 
    /// SttbfFfn structure:
    ///   fExtend  (2 bytes) — must be 0xFFFF for extended STTB
    ///   cData    (2 bytes) — number of font entries
    ///   cbExtra  (2 bytes) — extra data per entry (0 for fonts)
    ///   entries  — array of font definitions
    /// </summary>
    private void ReadFontTable()
    {
        if (_fib.FcSttbfFfn == 0 || _fib.LcbSttbfFfn == 0)
        {
            // No font table — use defaults
            AddDefaultFonts();
            return;
        }

        try
        {
            _tableReader.BaseStream.Seek(_fib.FcSttbfFfn, SeekOrigin.Begin);
            var endPos = _fib.FcSttbfFfn + _fib.LcbSttbfFfn;

            var header = _tableReader.ReadUInt16();
            int cData;
            if (header == 0xFFFF)
            {
                cData = _tableReader.ReadUInt16();
                _tableReader.ReadUInt16();
            }
            else
            {
                cData = header;
                _tableReader.ReadUInt16();
            }

            // Read each font entry (FFN structure)
            for (int i = 0; i < cData && _tableReader.BaseStream.Position < endPos; i++)
            {
                var font = ReadFfn(i);
                if (font != null)
                    Styles.Fonts.Add(font);
            }
        }
        catch
        {
            // If font table parsing fails, fall back to defaults
            if (Styles.Fonts.Count == 0)
                AddDefaultFonts();
        }
    }

    /// <summary>
    /// Reads a single FFN (Font Family Name) entry.
    /// 
    /// FFN structure:
    ///   cbFfnM1    (1 byte)  — total length of FFN minus 1
    ///   prq:2      (bits)    — pitch request (0=default, 1=fixed, 2=variable)
    ///   fTrueType:1(bit)     — is TrueType font
    ///   ff:3       (bits)    — font family
    ///   wWeight    (2 bytes) — font weight (0-1000)
    ///   chs        (1 byte)  — character set
    ///   ixchSzAlt  (1 byte)  — index into name for alternate name
    ///   panose     (10 bytes)— PANOSE classification
    ///   fs         (24 bytes)— FONTSIGNATURE
    ///   xszFfn     (variable)— UTF-16 font names, primary followed by alternate
    /// </summary>
    private FontDefinition? ReadFfn(int index)
    {
        var cbFfnM1 = _tableReader.ReadByte();
        var entryLength = cbFfnM1 + 1;
        if (entryLength < 40)
        {
            _tableReader.BaseStream.Seek(_tableReader.BaseStream.Position + cbFfnM1, SeekOrigin.Begin);
            return null;
        }

        var entryData = new byte[entryLength];
        entryData[0] = cbFfnM1;
        var remaining = _tableReader.Read(entryData, 1, cbFfnM1);
        if (remaining != cbFfnM1)
            return null;

        var packed = entryData[1];
        var prq = packed & 0x03;
        var fTrueType = (packed & 0x04) != 0;
        var ff = (packed >> 4) & 0x07;
        var chs = entryData[4];
        var ixchSzAlt = entryData[5];
        var fontNameChars = ReadFfnNameChars(entryData);
        var fontName = ReadNullTerminatedString(fontNameChars, 0);
        var altName = ixchSzAlt < fontNameChars.Length
            ? ReadNullTerminatedString(fontNameChars, ixchSzAlt)
            : null;

        return new FontDefinition
        {
            Index = index,
            Name = string.IsNullOrEmpty(fontName) ? $"Font{index}" : fontName,
            Family = ff,
            Charset = chs,
            Pitch = prq,
            Type = fTrueType ? 1 : 0,
            AltName = string.IsNullOrEmpty(altName) ? null : altName
        };
    }

    private static char[] ReadFfnNameChars(byte[] entryData)
    {
        const int nameOffset = 40;
        if (entryData.Length <= nameOffset)
            return Array.Empty<char>();

        var charCount = (entryData.Length - nameOffset) / 2;
        var chars = new char[charCount];
        for (int i = 0; i < charCount; i++)
        {
            chars[i] = (char)BinaryPrimitives.ReadUInt16LittleEndian(entryData.AsSpan(nameOffset + i * 2, 2));
        }

        return chars;
    }

    private static string ReadNullTerminatedString(char[] value, int startIndex)
    {
        if (startIndex < 0 || startIndex >= value.Length)
            return string.Empty;

        int endIndex = startIndex;
        while (endIndex < value.Length && value[endIndex] != '\0')
        {
            endIndex++;
        }

        return new string(value, startIndex, endIndex - startIndex);
    }

    private void AddDefaultFonts()
    {
        var defaultFonts = new[]
        {
            "Times New Roman", "Arial", "Courier New", "Symbol", "Wingdings",
            "Calibri", "Cambria", "宋体", "Tahoma", "Verdana"
        };

        for (int i = 0; i < defaultFonts.Length; i++)
        {
            Styles.Fonts.Add(new FontDefinition
            {
                Index = i,
                Name = defaultFonts[i],
                Family = i < 4 ? i : 1,
                Charset = 1,
                Pitch = 2,
                Type = 0
            });
        }
    }

    /// <summary>
    /// Reads style definitions from the STSH in the Table stream.
    /// Parses STD (Style Definition) entries from fcStshf/lcbStshf.
    /// </summary>
    private void ReadStyleDefinitions()
    {
        // First add default styles as fallback
        AddDefaultStyles();

        // Try to parse real styles from STSH
        if (_fib.FcStshf == 0 || _fib.LcbStshf == 0)
        {
            return;
        }

        try
        {
            ReadStsh();
            ResolveStyleInheritance();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Failed to read style sheet: {ex.Message}");
        }
    }

    /// <summary>
    /// Reads the STSH (Style Sheet) structure.
    ///
    /// STSH structure:
    ///   cstd (2 bytes) - count of styles
    ///   cbSTDBaseInFile (2 bytes) - base size of STD
    ///   flags (2 bytes) - flags
    ///   stiMaxWhenSaved (2 bytes) - max style index when saved
    ///   istdMaxFixedWhenSaved (2 bytes) - max fixed style index
    ///   nVerBuiltInNamesWhenSaved (2 bytes) - version of built-in names
    ///   cstd:1 (1 byte) - count of styles (if > 0x7FFF)
    ///   rgstd - array of STD entries
    /// </summary>
    private void ReadStsh()
    {
        _tableReader.BaseStream.Seek(_fib.FcStshf, SeekOrigin.Begin);
        var endPos = _fib.FcStshf + _fib.LcbStshf;

        // Read header
        var cstd = _tableReader.ReadUInt16();
        var cbSTDBaseInFile = _tableReader.ReadUInt16();
        var flags = _tableReader.ReadUInt16();
        var stiMaxWhenSaved = _tableReader.ReadUInt16();
        var istdMaxFixedWhenSaved = _tableReader.ReadUInt16();
        var nVerBuiltInNamesWhenSaved = _tableReader.ReadUInt16();

        // If cstd has high bit set, read extended count
        if ((cstd & 0x8000) != 0)
        {
            var cstdExtended = _tableReader.ReadUInt16();
            cstd = (ushort)((cstd & 0x7FFF) | (cstdExtended << 15));
        }

        // Read each STD entry
        for (int i = 0; i < cstd && _tableReader.BaseStream.Position < endPos; i++)
        {
            try
            {
                var style = ReadStd(i, cbSTDBaseInFile);
                if (style != null)
                {
                    // Replace or add style
                    var existing = Styles.Styles.FirstOrDefault(s => s.StyleId == style.StyleId);
                    if (existing != null)
                    {
                        Styles.Styles.Remove(existing);
                    }
                    Styles.Styles.Add(style);
                }
            }
            catch
            {
                // Skip malformed entries
                break;
            }
        }
    }

    /// <summary>
    /// Reads a single STD (Style Definition) entry.
    ///
    /// STD structure:
    ///   cbStd (2 bytes) - size of this STD
    ///   sti (12 bits) - style index (built-in style ID)
    ///   fScratch (1 bit) - scratch style
    ///   fInvalHeight (1 bit) - invalid height
    ///   fHasUpe (1 bit) - has up-to-date properties
    ///   fMassCopy (1 bit) - mass copied
    ///   sgc (4 bits) - style class (1=paragraph, 2=character, 3=table, 4=list)
    ///   istdBase (12 bits) - base style
    ///   cupx (4 bits) - count of UPX
    ///   istdNext (12 bits) - next style
    ///   bchUpe (2 bytes) - offset to UPE
    ///   fAutoRedef (1 bit) - auto redefine
    ///   fHidden (1 bit) - hidden style
    ///   f97LidsSet (1 bit) - lids set
    ///   fCopyLang (1 bit) - copy language
    ///   fPersonalCompose (1 bit) - personal compose
    ///   fPersonalReply (1 bit) - personal reply
    ///   fPersonal (1 bit) - personal
    ///   fNoHtmlDoc (1 bit) - no HTML document
    ///   fSemiHidden (1 bit) - semi-hidden
    ///   fLocked (1 bit) - locked
    ///   fInternalUse (1 bit) - internal use
    ///   xstzName (variable) - style name
    ///   rgupx (variable) - array of UPX (property exceptions)
    /// </summary>
    private StyleDefinition? ReadStd(int index, ushort cbSTDBase)
    {
        var startPos = _tableReader.BaseStream.Position;

        // Read cbStd
        var cbStd = _tableReader.ReadUInt16();
        if (cbStd == 0 || cbStd > 0x4000) // Sanity check
            return null;

        var entryEnd = startPos + 2 + cbStd;

        // Read first 2 bytes of STD base (packed bits)
        var word0 = _tableReader.ReadUInt16();
        var sti = word0 & 0x0FFF;
        var fScratch = (word0 >> 12) & 0x01;
        var fInvalHeight = (word0 >> 13) & 0x01;
        var fHasUpe = (word0 >> 14) & 0x01;
        var fMassCopy = (word0 >> 15) & 0x01;

        // Read second 2 bytes
        var word1 = _tableReader.ReadUInt16();
        var sgc = word1 & 0x000F;
        var istdBase = (word1 >> 4) & 0x0FFF;

        // Read third 2 bytes
        var word2 = _tableReader.ReadUInt16();
        var cupx = word2 & 0x000F;
        var istdNext = (word2 >> 4) & 0x0FFF;

        // Read bchUpe
        var bchUpe = _tableReader.ReadUInt16();

        // Read flags word
        var word4 = _tableReader.ReadUInt16();
        var fAutoRedef = (word4 >> 0) & 0x01;
        var fHidden = (word4 >> 1) & 0x01;
        var fSemiHidden = (word4 >> 9) & 0x01;
        var fLocked = (word4 >> 10) & 0x01;

        // Read style name (XSTZ - counted string with null terminator)
        var nameLength = _tableReader.ReadUInt16();
        if (nameLength > 64) // Sanity check
            nameLength = 64;

        string styleName;
        if (nameLength > 0)
        {
            var nameBytes = _tableReader.ReadBytes(nameLength * 2); // Unicode
            styleName = Encoding.Unicode.GetString(nameBytes);
        }
        else
        {
            styleName = $"Style{index}";
        }

        // Skip null terminator if present
        if (_tableReader.ReadByte() != 0)
            _tableReader.BaseStream.Seek(-1, SeekOrigin.Current);

        // Determine style type from sgc
        var styleType = sgc switch
        {
            1 => StyleType.Paragraph,
            2 => StyleType.Character,
            3 => StyleType.Table,
            4 => StyleType.Numbering,
            _ => StyleType.Paragraph
        };

        // Create style definition
        var style = new StyleDefinition
        {
            StyleId = (ushort)index,
            Name = styleName,
            Type = styleType,
            BasedOn = istdBase < 0x0FFF ? (ushort?)istdBase : null,
            NextParagraphStyle = istdNext < 0x0FFF ? (ushort?)istdNext : null,
            IsHidden = fHidden != 0,
            IsAutoRedefined = fAutoRedef != 0,
            IsLinked = false,
            IsPrimary = sti < 266,
            TableProperties = styleType == StyleType.Table ? new TableProperties() : null,
            ParagraphProperties = new ParagraphProperties(),
            RunProperties = new RunProperties()
        };

        var sprmParser = new SprmParser(_tableReader, 0);
        var fkpParser = new FkpParser(null!, null!, _fib, null!);

        // Read properties (UPX)
        // For paragraph styles, first UPX is PAP, second is CHP
        // For character styles, first UPX is CHP
        // Each UPX starts with a 2-byte size
        for (int i = 0; i < cupx; i++)
        {
            if (_tableReader.BaseStream.Position + 2 > entryEnd) break;
            
            var cbUpx = _tableReader.ReadUInt16();
            if (cbUpx == 0) continue;
            
            var upxEnd = _tableReader.BaseStream.Position + cbUpx;
            if (upxEnd > entryEnd) break;

            var grpprl = _tableReader.ReadBytes(cbUpx);
            
            if (styleType == StyleType.Paragraph)
            {
                if (i == 0) // PAP UPX
                {
                    // First 2 bytes of PAP UPX is istd (style index)
                    if (cbUpx > 2)
                    {
                        var papGrpprl = new byte[cbUpx - 2];
                        Array.Copy(grpprl, 2, papGrpprl, 0, cbUpx - 2);
                        var pap = new PapBase();
                        sprmParser.ApplyToPap(papGrpprl, pap);
                        style.ParagraphProperties = fkpParser.ConvertToParagraphProperties(pap, Styles);
                    }
                }
                else if (i == 1) // CHP UPX — per MS-DOC may start with 2-byte istd (char style ref); skip so grpprl is correct
                {
                    var chp = new ChpBase();
                    byte[] chpGrpprl;
                    if (cbUpx > 2)
                    {
                        chpGrpprl = new byte[cbUpx - 2];
                        Array.Copy(grpprl, 2, chpGrpprl, 0, chpGrpprl.Length);
                    }
                    else
                        chpGrpprl = grpprl;
                    sprmParser.ApplyToChp(chpGrpprl, chp);
                    style.RunProperties = fkpParser.ConvertToRunProperties(chp, Styles);
                }
            }
            else if (styleType == StyleType.Character)
            {
                if (i == 0) // CHP UPX
                {
                    var chp = new ChpBase();
                    sprmParser.ApplyToChp(grpprl, chp);
                    style.RunProperties = fkpParser.ConvertToRunProperties(chp, Styles);
                }
            }
            else if (styleType == StyleType.Table)
            {
                if (i == 0)
                {
                    var tap = new TapBase();
                    sprmParser.ApplyToTap(grpprl, tap);
                    style.TableProperties = ConvertToTableProperties(tap);
                }
                else if (i == 1)
                {
                    var pap = new PapBase();
                    sprmParser.ApplyToPap(SkipStylePrefix(grpprl), pap);
                    style.ParagraphProperties = fkpParser.ConvertToParagraphProperties(pap, Styles);
                }
                else if (i == 2)
                {
                    var chp = new ChpBase();
                    sprmParser.ApplyToChp(SkipStylePrefix(grpprl), chp);
                    style.RunProperties = fkpParser.ConvertToRunProperties(chp, Styles);
                }
            }
            
            _tableReader.BaseStream.Seek(upxEnd, SeekOrigin.Begin);
            if (cbUpx % 2 != 0) _tableReader.BaseStream.Seek(1, SeekOrigin.Current); // 2-byte alignment
        }

        // Skip to end of entry
        _tableReader.BaseStream.Seek(entryEnd, SeekOrigin.Begin);

        return style;
    }

    /// <summary>
    /// Recursively resolves style inheritance by merging properties from base styles.
    /// </summary>
    private void ResolveStyleInheritance()
    {
        var resolved = new HashSet<ushort>();

        foreach (var style in Styles.Styles)
        {
            ResolveStyle(style, resolved, new HashSet<ushort>());
        }
    }

    private void ResolveStyle(StyleDefinition style, HashSet<ushort> resolved, HashSet<ushort> visiting)
    {
        if (resolved.Contains(style.StyleId)) return;
        if (!style.BasedOn.HasValue)
        {
            resolved.Add(style.StyleId);
            return;
        }

        // Detect circular dependency
        if (visiting.Contains(style.StyleId))
        {
            Logger.Warning($"Circular style inheritance detected for style ID {style.StyleId}");
            resolved.Add(style.StyleId);
            return;
        }

        visiting.Add(style.StyleId);

        var baseStyle = GetStyle(style.BasedOn.Value);
        if (baseStyle != null)
        {
            // Ensure base style is resolved first
            ResolveStyle(baseStyle, resolved, visiting);

            if (baseStyle.TableProperties != null)
            {
                style.TableProperties ??= new TableProperties();
                style.TableProperties.MergeWith(baseStyle.TableProperties);
            }

            // Merge properties from base style
            if (baseStyle.ParagraphProperties != null)
            {
                style.ParagraphProperties ??= new ParagraphProperties();
                style.ParagraphProperties.MergeWith(baseStyle.ParagraphProperties);
            }

            if (baseStyle.RunProperties != null)
            {
                style.RunProperties ??= new RunProperties();
                style.RunProperties.MergeWith(baseStyle.RunProperties);
            }
        }

        visiting.Remove(style.StyleId);
        resolved.Add(style.StyleId);
    }

    private static byte[] SkipStylePrefix(byte[] grpprl)
    {
        if (grpprl.Length <= 2)
            return grpprl;

        var trimmed = new byte[grpprl.Length - 2];
        Array.Copy(grpprl, 2, trimmed, 0, trimmed.Length);
        return trimmed;
    }

    private static TableProperties ConvertToTableProperties(TapBase tap)
    {
        return new TableProperties
        {
            Alignment = tap.Justification switch
            {
                1 => TableAlignment.Center,
                2 => TableAlignment.Right,
                _ => TableAlignment.Left
            },
            CellSpacing = tap.CellSpacing != 0 ? tap.CellSpacing : (tap.GapHalf != 0 ? tap.GapHalf * 2 : 0),
            Indent = tap.IndentLeft,
            PreferredWidth = tap.TableWidth,
            BorderTop = tap.BorderTop,
            BorderBottom = tap.BorderBottom,
            BorderLeft = tap.BorderLeft,
            BorderRight = tap.BorderRight,
            BorderInsideH = tap.BorderInsideH,
            BorderInsideV = tap.BorderInsideV,
            Shading = tap.Shading
        };
    }

    /// <summary>
    /// Adds default built-in styles.
    /// </summary>
    private void AddDefaultStyles()
    {
        // Normal style
        Styles.Styles.Add(new StyleDefinition
        {
            StyleId = 0,
            Name = "Normal",
            Type = StyleType.Paragraph,
            ParagraphProperties = new ParagraphProperties(),
            RunProperties = new RunProperties { FontSize = 24 }
        });

        // Heading styles 1-9
        for (int i = 0; i < 9; i++)
        {
            Styles.Styles.Add(new StyleDefinition
            {
                StyleId = (ushort)(i + 1),
                Name = $"heading {i + 1}",
                Type = StyleType.Paragraph,
                BasedOn = 0,
                ParagraphProperties = new ParagraphProperties
                {
                    SpaceBefore = (9 - i) * 60,
                    SpaceAfter = 60
                },
                RunProperties = new RunProperties
                {
                    FontSize = (ushort)((16 - i * 2) * 2),
                    IsBold = true
                }
            });
        }

        // Default Paragraph Font (character style)
        Styles.Styles.Add(new StyleDefinition
        {
            StyleId = 10,
            Name = "Default Paragraph Font",
            Type = StyleType.Character,
            RunProperties = new RunProperties { FontSize = 24 }
        });
    }

    /// <summary>
    /// Gets font name by index.
    /// </summary>
    public string? GetFontName(int index)
    {
        if (index < 0 || index >= Styles.Fonts.Count) return null;
        return Styles.Fonts[index].Name;
    }

    /// <summary>
    /// Gets style by ID.
    /// </summary>
    public StyleDefinition? GetStyle(ushort styleId)
    {
        return Styles.Styles.FirstOrDefault(s => s.StyleId == styleId);
    }

    /// <summary>
    /// Gets style by name (case-insensitive).
    /// </summary>
    public StyleDefinition? GetStyleByName(string name)
    {
        return Styles.Styles.FirstOrDefault(s =>
            s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }
}
