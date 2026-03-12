using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Writers;

public class NumberingWriter
{
    private readonly XmlWriter _writer;
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private DocumentModel? _document;

    private sealed class NumberingInstanceDefinition
    {
        public int NumId { get; init; }
        public int AbstractNumId { get; init; }
        public NumberingDefinition? BaseDefinition { get; init; }
        public List<ListLevelOverride> LevelOverrides { get; init; } = new();
    }

    private sealed class NumberingPackage
    {
        public List<NumberingDefinition> AbstractDefinitions { get; init; } = new();
        public List<NumberingInstanceDefinition> Instances { get; init; } = new();
    }

    public NumberingWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    public void WriteNumbering(DocumentModel document)
    {
        _document = document;
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "numbering", WNs);
        _writer.WriteAttributeString("xmlns", "w", null, WNs);
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var numberingPackage = BuildNumberingDefinitions(document);

        if (numberingPackage.AbstractDefinitions.Count > 0)
        {
            foreach (var numDef in numberingPackage.AbstractDefinitions)
            {
                var id = numDef.Id;
                if (id <= 0)
                {
                    continue;
                }

                WriteAbstractNum(numDef, id);
            }

            foreach (var instance in numberingPackage.Instances)
            {
                WriteNum(instance);
            }
        }
        else
        {
            WriteDefaultNumbering();
        }

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
        _document = null;
    }

    private static NumberingPackage BuildNumberingDefinitions(DocumentModel document)
    {
        var definitions = document.NumberingDefinitions
            .Where(definition => definition.Id > 0)
            .GroupBy(definition => definition.Id)
            .Select(group => group.First())
            .ToDictionary(definition => definition.Id);

        var usedListIds = document.Paragraphs
            .Select(paragraph => (paragraph.Properties?.ListFormatId ?? 0) > 0
                ? paragraph.Properties!.ListFormatId
                : paragraph.ListFormatId)
            .Where(listId => listId > 0)
            .Distinct()
            .OrderBy(listId => listId)
            .ToList();

        var overrides = document.ListFormatOverrides
            .Where(overrideDefinition => overrideDefinition.OverrideId > 0)
            .GroupBy(overrideDefinition => overrideDefinition.OverrideId)
            .Select(group => group.First())
            .ToDictionary(overrideDefinition => overrideDefinition.OverrideId);

        foreach (var listOverride in overrides.Values)
        {
            var abstractNumId = listOverride.ListId > 0 ? listOverride.ListId : listOverride.OverrideId;
            if (!definitions.ContainsKey(abstractNumId))
            {
                definitions[abstractNumId] = CreateFallbackDefinition(abstractNumId);
            }
        }

        foreach (var listId in usedListIds)
        {
            var abstractNumId = overrides.TryGetValue(listId, out var listOverride) && listOverride.ListId > 0
                ? listOverride.ListId
                : listId;

            if (definitions.ContainsKey(abstractNumId))
            {
                continue;
            }

            definitions[abstractNumId] = CreateFallbackDefinition(abstractNumId);
        }

        var instanceIds = usedListIds.Count > 0
            ? usedListIds
            : overrides.Keys.Union(definitions.Keys).OrderBy(id => id).ToList();

        var instances = new List<NumberingInstanceDefinition>();
        foreach (var numId in instanceIds)
        {
            ListFormatOverride? listOverride = null;
            var abstractNumId = overrides.TryGetValue(numId, out listOverride) && listOverride.ListId > 0
                ? listOverride.ListId
                : numId;

            if (!definitions.ContainsKey(abstractNumId))
            {
                definitions[abstractNumId] = CreateFallbackDefinition(abstractNumId);
            }

            instances.Add(new NumberingInstanceDefinition
            {
                NumId = numId,
                AbstractNumId = abstractNumId,
                BaseDefinition = definitions[abstractNumId],
                LevelOverrides = listOverride != null
                    ? GetEffectiveLevelOverrides(definitions[abstractNumId], listOverride)
                    : new List<ListLevelOverride>()
            });
        }

        return new NumberingPackage
        {
            AbstractDefinitions = definitions.Values.OrderBy(definition => definition.Id).ToList(),
            Instances = instances.OrderBy(instance => instance.NumId).ToList()
        };
    }

    private static List<ListLevelOverride> GetEffectiveLevelOverrides(NumberingDefinition definition, ListFormatOverride listOverride)
    {
        var baseStarts = definition.Levels.ToDictionary(level => level.Level, level => level.Start);
        var effectiveOverrides = new List<ListLevelOverride>();

        foreach (var levelOverride in listOverride.Levels.OrderBy(level => level.Level))
        {
            bool hasEffectiveStartOverride = levelOverride.HasStartAt && levelOverride.StartAt > 0;
            if (hasEffectiveStartOverride && baseStarts.TryGetValue(levelOverride.Level, out var baseStart) && baseStart == levelOverride.StartAt)
            {
                hasEffectiveStartOverride = false;
            }

            if (!hasEffectiveStartOverride && !levelOverride.HasFormattingOverride)
            {
                continue;
            }

            effectiveOverrides.Add(new ListLevelOverride
            {
                Level = levelOverride.Level,
                StartAt = levelOverride.StartAt,
                HasStartAt = hasEffectiveStartOverride,
                HasFormattingOverride = levelOverride.HasFormattingOverride,
                Alignment = levelOverride.Alignment,
                NumberFormat = levelOverride.NumberFormat,
                NumberText = levelOverride.NumberText,
                ParagraphProperties = levelOverride.ParagraphProperties,
                RunProperties = levelOverride.RunProperties
            });
        }

        return effectiveOverrides;
    }

    private static NumberingDefinition CreateFallbackDefinition(int listId)
    {
        var definition = new NumberingDefinition
        {
            Id = listId
        };

        for (int level = 0; level < 9; level++)
        {
            definition.Levels.Add(new NumberingLevel
            {
                Level = level,
                NumberFormat = NumberFormat.Decimal,
                Text = $"%{level + 1}.",
                Start = 1
            });
        }

        return definition;
    }

    private void WriteDefaultNumbering()
    {
        _writer.WriteStartElement("w", "abstractNum", WNs);
        _writer.WriteAttributeString("w", "abstractNumId", WNs, "0");

        _writer.WriteStartElement("w", "nsid", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "00000000");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "multiLevelType", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "hybridMultilevel");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "tmpl", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "00000000");
        _writer.WriteEndElement();

        for (int lvl = 0; lvl < 9; lvl++)
        {
            WriteLevel(new NumberingLevel
            {
                Level = lvl,
                NumberFormat = NumberFormat.Decimal,
                Text = $"%{lvl + 1}.",
                Start = 1
            });
        }

        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "num", WNs);
        _writer.WriteAttributeString("w", "numId", WNs, "1");
        _writer.WriteStartElement("w", "abstractNumId", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "0");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }

    private void WriteAbstractNum(NumberingDefinition numDef, int abstractNumId)
    {
        _writer.WriteStartElement("w", "abstractNum", WNs);
        _writer.WriteAttributeString("w", "abstractNumId", WNs, abstractNumId.ToString());

        _writer.WriteStartElement("w", "nsid", WNs);
        _writer.WriteAttributeString("w", "val", WNs, Convert.ToString(numDef.Id, 16).PadLeft(8, '0').ToUpper());
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "multiLevelType", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "hybridMultilevel");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "tmpl", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "00000000");
        _writer.WriteEndElement();

        // Write levels from actual definition data
        if (numDef.Levels.Count > 0)
        {
            foreach (var level in numDef.Levels)
            {
                WriteLevel(level);
            }
        }
        else
        {
            // Fallback: write a single bullet level
            WriteLevel(new NumberingLevel
            {
                Level = 0,
                NumberFormat = NumberFormat.Bullet,
                Text = "\u00B7",
                Start = 1
            });
        }

        _writer.WriteEndElement(); // w:abstractNum
    }

    private void WriteLevel(NumberingLevel level)
    {
        _writer.WriteStartElement("w", "lvl", WNs);
        _writer.WriteAttributeString("w", "ilvl", WNs, level.Level.ToString());

        _writer.WriteStartElement("w", "start", WNs);
        _writer.WriteAttributeString("w", "val", WNs, level.Start.ToString());
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "numFmt", WNs);
        _writer.WriteAttributeString("w", "val", WNs, GetNumberFormatValue(level.NumberFormat));
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlText", WNs);
        _writer.WriteAttributeString("w", "val", WNs, level.Text ?? $"%{level.Level + 1}.");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlJc", WNs);
        _writer.WriteAttributeString("w", "val", WNs, GetLevelAlignmentValue(level.Alignment));
        _writer.WriteEndElement();

        WriteLevelParagraphProperties(level);

        if (level.RunProperties != null && RunPropertiesHelper.HasRunProperties(level.RunProperties))
        {
            RunPropertiesHelper.WriteStyleRunProperties(_writer, level.RunProperties, _document?.Theme);
        }

        _writer.WriteEndElement(); // w:lvl
    }

    private void WriteNum(NumberingInstanceDefinition instance)
    {
        _writer.WriteStartElement("w", "num", WNs);
        _writer.WriteAttributeString("w", "numId", WNs, instance.NumId.ToString());

        _writer.WriteStartElement("w", "abstractNumId", WNs);
        _writer.WriteAttributeString("w", "val", WNs, instance.AbstractNumId.ToString());
        _writer.WriteEndElement();

        foreach (var levelOverride in instance.LevelOverrides)
        {
            _writer.WriteStartElement("w", "lvlOverride", WNs);
            _writer.WriteAttributeString("w", "ilvl", WNs, levelOverride.Level.ToString());

            if (levelOverride.HasStartAt)
            {
                _writer.WriteStartElement("w", "startOverride", WNs);
                _writer.WriteAttributeString("w", "val", WNs, levelOverride.StartAt.ToString());
                _writer.WriteEndElement();
            }

            if (levelOverride.HasFormattingOverride)
            {
                WriteLevel(BuildOverrideLevel(instance.BaseDefinition, levelOverride));
            }

            _writer.WriteEndElement();
        }

        _writer.WriteEndElement();
    }

    private string GetNumberFormatValue(NumberFormat format)
    {
        return format switch
        {
            NumberFormat.Bullet => "bullet",
            NumberFormat.Decimal => "decimal",
            NumberFormat.LowerRoman => "lowerRoman",
            NumberFormat.UpperRoman => "upperRoman",
            NumberFormat.LowerLetter => "lowerLetter",
            NumberFormat.UpperLetter => "upperLetter",
            NumberFormat.Ordinal => "ordinal",
            NumberFormat.CardinalText => "cardinalText",
            NumberFormat.OrdinalText => "ordinalText",
            NumberFormat.Hex => "hex",
            NumberFormat.Chicago => "chicago",
            NumberFormat.IdeographDigital => "ideographDigital",
            NumberFormat.JapaneseCounting => "japaneseCounting",
            NumberFormat.Aiueo => "aiueo",
            NumberFormat.Iroha => "iroha",
            NumberFormat.DecimalFullWidth => "decimalFullWidth",
            NumberFormat.DecimalHalfWidth => "decimalHalfWidth",
            NumberFormat.JapaneseLegal => "japaneseLegal",
            NumberFormat.JapaneseDigitalTenThousand => "japaneseDigitalTenThousand",
            NumberFormat.DecimalEnclosedCircle => "decimalEnclosedCircle",
            NumberFormat.AiueoFullWidth => "aiueoFullWidth",
            NumberFormat.IrohaFullWidth => "irohaFullWidth",
            NumberFormat.DecimalZero => "decimalZero",
            NumberFormat.Ganada => "ganada",
            NumberFormat.Chosung => "chosung",
            NumberFormat.DecimalEnclosedFullstop => "decimalEnclosedFullstop",
            NumberFormat.DecimalEnclosedParen => "decimalEnclosedParen",
            NumberFormat.DecimalEnclosedCircleChinese => "decimalEnclosedCircleChinese",
            NumberFormat.IdeographEnclosedCircle => "ideographEnclosedCircle",
            NumberFormat.IdeographTraditional => "ideographTraditional",
            NumberFormat.IdeographZodiac => "ideographZodiac",
            NumberFormat.IdeographZodiacTraditional => "ideographZodiacTraditional",
            NumberFormat.TaiwaneseCounting => "taiwaneseCounting",
            NumberFormat.IdeographLegalTraditional => "ideographLegalTraditional",
            NumberFormat.TaiwaneseCountingThousand => "taiwaneseCountingThousand",
            NumberFormat.TaiwaneseDigital => "taiwaneseDigital",
            NumberFormat.ChineseCounting => "chineseCounting",
            NumberFormat.ChineseLegalSimplified => "chineseLegalSimplified",
            NumberFormat.ChineseLegalTraditional => "chineseLegalTraditional",
            NumberFormat.JapaneseCounting2 => "japaneseCounting2",
            NumberFormat.JapaneseDigitalHundredCount => "japaneseDigitalHundredCount",
            NumberFormat.JapaneseDigitalThousandCount => "japaneseDigitalThousandCount",
            _ => "decimal"
        };
    }

    private static NumberingLevel BuildOverrideLevel(NumberingDefinition? baseDefinition, ListLevelOverride levelOverride)
    {
        var baseLevel = baseDefinition?.Levels.FirstOrDefault(level => level.Level == levelOverride.Level);

        return new NumberingLevel
        {
            Level = levelOverride.Level,
            Start = levelOverride.HasStartAt && levelOverride.StartAt > 0
                ? levelOverride.StartAt
                : baseLevel?.Start ?? 1,
            Alignment = levelOverride.HasFormattingOverride
                ? levelOverride.Alignment
                : baseLevel?.Alignment ?? 0,
            NumberFormat = levelOverride.NumberFormat ?? baseLevel?.NumberFormat ?? NumberFormat.Decimal,
            Text = levelOverride.NumberText ?? baseLevel?.Text ?? $"%{levelOverride.Level + 1}.",
            ParagraphProperties = levelOverride.ParagraphProperties ?? baseLevel?.ParagraphProperties,
            RunProperties = levelOverride.RunProperties ?? baseLevel?.RunProperties
        };
    }

    private static string GetLevelAlignmentValue(int alignment)
    {
        return alignment switch
        {
            1 => "center",
            2 => "right",
            _ => "left"
        };
    }

    private void WriteLevelParagraphProperties(NumberingLevel level)
    {
        var props = level.ParagraphProperties;
        int defaultIndentLeft = 720 + level.Level * 720;
        int defaultHanging = 360;
        bool hasParagraphFormatting = props != null &&
            (props.KeepWithNext || props.KeepTogether || props.PageBreakBefore ||
             props.BorderTop != null || props.BorderBottom != null || props.BorderLeft != null || props.BorderRight != null ||
             props.Shading != null ||
             props.SpaceBefore > 0 || props.SpaceBeforeLines > 0 || props.SpaceAfter > 0 || props.SpaceAfterLines > 0 ||
             props.HasExplicitLineSpacing || props.LineSpacing != 240 || props.LineSpacingMultiple != 1 ||
             props.IndentLeft != 0 || props.IndentLeftChars != 0 || props.IndentRight != 0 || props.IndentRightChars != 0 ||
             props.IndentFirstLine != 0 || props.IndentFirstLineChars != 0 ||
             props.Alignment != ParagraphAlignment.Left ||
             !props.WordWrap || !props.Kinsoku || !props.SnapToGrid || !props.AutoSpaceDe || !props.AutoSpaceDn ||
             props.TopLinePunct || props.OverflowPunct ||
             (props.OutlineLevel >= 0 && props.OutlineLevel < 9));

        _writer.WriteStartElement("w", "pPr", WNs);

        if (props?.KeepWithNext == true)
        {
            _writer.WriteStartElement("w", "keepNext", WNs);
            _writer.WriteEndElement();
        }

        if (props?.KeepTogether == true)
        {
            _writer.WriteStartElement("w", "keepLines", WNs);
            _writer.WriteEndElement();
        }

        if (props?.PageBreakBefore == true)
        {
            _writer.WriteStartElement("w", "pageBreakBefore", WNs);
            _writer.WriteEndElement();
        }

        if (props != null && (props.BorderTop != null || props.BorderBottom != null || props.BorderLeft != null || props.BorderRight != null))
        {
            _writer.WriteStartElement("w", "pBdr", WNs);
            if (props.BorderTop != null) WriteBorder("top", props.BorderTop);
            if (props.BorderBottom != null) WriteBorder("bottom", props.BorderBottom);
            if (props.BorderLeft != null) WriteBorder("left", props.BorderLeft);
            if (props.BorderRight != null) WriteBorder("right", props.BorderRight);
            _writer.WriteEndElement();
        }

        if (props?.Shading != null)
        {
            WriteShading(props.Shading);
        }

        bool hasExplicitLineSpacing = props != null && (props.HasExplicitLineSpacing || props.LineSpacing != 240 || props.LineSpacingMultiple != 1);
        if (props != null && (props.SpaceBefore > 0 || props.SpaceBeforeLines > 0 || props.SpaceAfter > 0 || props.SpaceAfterLines > 0 || hasExplicitLineSpacing))
        {
            _writer.WriteStartElement("w", "spacing", WNs);
            if (props.SpaceBeforeLines > 0)
                _writer.WriteAttributeString("w", "beforeLines", WNs, props.SpaceBeforeLines.ToString());
            else if (props.SpaceBefore > 0)
                _writer.WriteAttributeString("w", "before", WNs, props.SpaceBefore.ToString());

            if (props.SpaceAfterLines > 0)
                _writer.WriteAttributeString("w", "afterLines", WNs, props.SpaceAfterLines.ToString());
            else if (props.SpaceAfter > 0)
                _writer.WriteAttributeString("w", "after", WNs, props.SpaceAfter.ToString());

            if (hasExplicitLineSpacing)
            {
                int lineVal = props.LineSpacing;
                string lineRule;
                if (props.LineSpacingMultiple == 1)
                {
                    lineRule = "auto";
                }
                else if (lineVal < 0)
                {
                    lineVal = Math.Abs(lineVal);
                    lineRule = "exact";
                }
                else
                {
                    lineRule = "atLeast";
                }

                _writer.WriteAttributeString("w", "line", WNs, lineVal.ToString());
                _writer.WriteAttributeString("w", "lineRule", WNs, lineRule);
            }

            _writer.WriteEndElement();
        }

        WriteIndentation(props, defaultIndentLeft, defaultHanging, level.RunProperties);

        if (props != null && props.Alignment != ParagraphAlignment.Left)
        {
            _writer.WriteStartElement("w", "jc", WNs);
            _writer.WriteAttributeString("w", "val", WNs, GetParagraphAlignmentValue(props.Alignment));
            _writer.WriteEndElement();
        }

        if (props != null && props.OutlineLevel >= 0 && props.OutlineLevel < 9)
        {
            _writer.WriteStartElement("w", "outlineLvl", WNs);
            _writer.WriteAttributeString("w", "val", WNs, props.OutlineLevel.ToString());
            _writer.WriteEndElement();
        }

        if (props != null && !props.WordWrap)
        {
            _writer.WriteStartElement("w", "wordWrap", WNs);
            _writer.WriteAttributeString("w", "val", WNs, "0");
            _writer.WriteEndElement();
        }

        if (props != null && !props.Kinsoku)
        {
            _writer.WriteStartElement("w", "kinsoku", WNs);
            _writer.WriteAttributeString("w", "val", WNs, "0");
            _writer.WriteEndElement();
        }

        if (props != null && !props.SnapToGrid)
        {
            _writer.WriteStartElement("w", "snapToGrid", WNs);
            _writer.WriteAttributeString("w", "val", WNs, "0");
            _writer.WriteEndElement();
        }

        if (props != null && !props.AutoSpaceDe)
        {
            _writer.WriteStartElement("w", "autoSpaceDE", WNs);
            _writer.WriteAttributeString("w", "val", WNs, "0");
            _writer.WriteEndElement();
        }

        if (props != null && !props.AutoSpaceDn)
        {
            _writer.WriteStartElement("w", "autoSpaceDN", WNs);
            _writer.WriteAttributeString("w", "val", WNs, "0");
            _writer.WriteEndElement();
        }

        if (props != null && props.TopLinePunct)
        {
            _writer.WriteStartElement("w", "topLinePunct", WNs);
            _writer.WriteEndElement();
        }

        if (props != null && props.OverflowPunct)
        {
            _writer.WriteStartElement("w", "overflowPunct", WNs);
            _writer.WriteEndElement();
        }

        if (!hasParagraphFormatting)
        {
            // Preserve the previous default numbering indent behaviour even when no explicit pPr exists.
            // The indentation was already emitted by WriteIndentation above.
        }

        _writer.WriteEndElement();
    }

    private void WriteIndentation(ParagraphProperties? props, int defaultIndentLeft, int defaultHanging, RunProperties? runProperties)
    {
        int fontSizeHalfPoints = runProperties?.FontSize > 0 ? runProperties.FontSize : 24;
        int indentLeft = defaultIndentLeft;
        if (props != null)
        {
            if (props.IndentLeft != 0)
            {
                indentLeft = props.IndentLeft;
            }
            else if (props.IndentLeftChars != 0)
            {
                indentLeft = ConvertCharacterIndentToTwips(props.IndentLeftChars, fontSizeHalfPoints);
            }
        }

        int? indentRight = null;
        if (props != null)
        {
            indentRight = props.IndentRight != 0
                ? props.IndentRight
                : props.IndentRightChars != 0
                    ? ConvertCharacterIndentToTwips(props.IndentRightChars, fontSizeHalfPoints)
                    : null;
        }

        int? firstLine = null;
        int? hanging = null;
        if (props != null)
        {
            if (props.IndentFirstLineChars > 0)
            {
                firstLine = props.IndentFirstLine > 0
                    ? props.IndentFirstLine
                    : ConvertCharacterIndentToTwips(props.IndentFirstLineChars, fontSizeHalfPoints);
            }
            else if (props.IndentFirstLineChars < 0)
            {
                hanging = props.IndentFirstLine < 0
                    ? Math.Abs(props.IndentFirstLine)
                    : ConvertCharacterIndentToTwips(props.IndentFirstLineChars, fontSizeHalfPoints);
            }
            else if (props.IndentFirstLine > 0)
            {
                firstLine = props.IndentFirstLine;
            }
            else if (props.IndentFirstLine < 0)
            {
                hanging = Math.Abs(props.IndentFirstLine);
            }
        }

        hanging ??= defaultHanging;

        _writer.WriteStartElement("w", "ind", WNs);
        _writer.WriteAttributeString("w", "left", WNs, indentLeft.ToString());
        if (props != null && props.IndentLeftChars != 0)
            _writer.WriteAttributeString("w", "leftChars", WNs, props.IndentLeftChars.ToString());
        if (indentRight.HasValue)
            _writer.WriteAttributeString("w", "right", WNs, indentRight.Value.ToString());
        if (props != null && props.IndentRightChars != 0)
            _writer.WriteAttributeString("w", "rightChars", WNs, props.IndentRightChars.ToString());
        if (firstLine.HasValue && firstLine.Value > 0)
            _writer.WriteAttributeString("w", "firstLine", WNs, firstLine.Value.ToString());
        if (props != null && props.IndentFirstLineChars > 0)
            _writer.WriteAttributeString("w", "firstLineChars", WNs, props.IndentFirstLineChars.ToString());
        if (hanging.HasValue && hanging.Value > 0)
            _writer.WriteAttributeString("w", "hanging", WNs, hanging.Value.ToString());
        if (props != null && props.IndentFirstLineChars < 0)
            _writer.WriteAttributeString("w", "hangingChars", WNs, Math.Abs(props.IndentFirstLineChars).ToString());
        _writer.WriteEndElement();
    }

    private void WriteBorder(string position, BorderInfo border)
    {
        if (border.Style == BorderStyle.None || IsLikelyMalformedBorder(border))
            return;

        string? themeColor = ColorHelper.GetThemeColorName(border.Color);
        string? resolvedThemeHex = ColorHelper.ResolveThemeColorHex(border.Color, _document?.Theme);
        _writer.WriteStartElement("w", position, WNs);
        _writer.WriteAttributeString("w", "val", WNs, GetBorderStyle(border.Style));
        _writer.WriteAttributeString("w", "sz", WNs, border.Width.ToString());
        _writer.WriteAttributeString("w", "space", WNs, border.Space.ToString());
        _writer.WriteAttributeString("w", "color", WNs, resolvedThemeHex ?? ColorHelper.ResolveColorHex(border.Color, _document?.Theme));
        if (themeColor != null)
            _writer.WriteAttributeString("w", "themeColor", WNs, themeColor);
        _writer.WriteEndElement();
    }

    private void WriteShading(ShadingInfo shading)
    {
        string? foregroundThemeColor = ColorHelper.GetThemeColorName(shading.ForegroundColor);
        string? foregroundThemeHex = ColorHelper.ResolveThemeColorHex(shading.ForegroundColor, _document?.Theme);
        string? backgroundThemeColor = ColorHelper.GetThemeColorName(shading.BackgroundColor);
        string? backgroundThemeHex = ColorHelper.ResolveThemeColorHex(shading.BackgroundColor, _document?.Theme);
        _writer.WriteStartElement("w", "shd", WNs);
        _writer.WriteAttributeString("w", "val", WNs, !string.IsNullOrEmpty(shading.PatternVal) ? shading.PatternVal : ShadingPatternToShdVal(shading.Pattern));
        if (shading.ForegroundColor != 0)
        {
            _writer.WriteAttributeString("w", "color", WNs, foregroundThemeHex ?? ColorHelper.ResolveColorHex(shading.ForegroundColor, _document?.Theme));
            if (foregroundThemeColor != null)
                _writer.WriteAttributeString("w", "themeColor", WNs, foregroundThemeColor);
        }
        _writer.WriteAttributeString("w", "fill", WNs, backgroundThemeHex ?? ColorHelper.ResolveColorHex(shading.BackgroundColor, _document?.Theme, fallback: "FFFFFF"));
        if (backgroundThemeColor != null)
            _writer.WriteAttributeString("w", "themeFill", WNs, backgroundThemeColor);
        _writer.WriteEndElement();
    }

    private static string ShadingPatternToShdVal(ShadingPattern pattern)
    {
        return pattern switch
        {
            ShadingPattern.Clear => "clear",
            ShadingPattern.Solid => "solid",
            ShadingPattern.Percent5 => "pct5",
            ShadingPattern.Percent10 => "pct10",
            ShadingPattern.Percent20 => "pct20",
            ShadingPattern.Percent25 => "pct25",
            ShadingPattern.Percent30 => "pct30",
            ShadingPattern.Percent40 => "pct40",
            ShadingPattern.Percent50 => "pct50",
            ShadingPattern.Percent60 => "pct60",
            ShadingPattern.Percent70 => "pct70",
            ShadingPattern.Percent75 => "pct75",
            ShadingPattern.Percent80 => "pct80",
            ShadingPattern.Percent90 => "pct90",
            ShadingPattern.LightHorizontal => "thinHorzStripe",
            ShadingPattern.DarkHorizontal => "horzStripe",
            ShadingPattern.LightVertical => "thinVertStripe",
            ShadingPattern.DarkVertical => "vertStripe",
            ShadingPattern.LightDiagonalDown => "thinDiagStripe",
            ShadingPattern.LightDiagonalUp => "thinReverseDiagStripe",
            ShadingPattern.DarkDiagonalDown => "diagStripe",
            ShadingPattern.DarkDiagonalUp => "reverseDiagStripe",
            ShadingPattern.DarkGrid => "horzCross",
            ShadingPattern.DarkTrellis => "diagCross",
            ShadingPattern.LightGray => "pct25",
            ShadingPattern.MediumGray => "pct50",
            ShadingPattern.DarkGray => "pct75",
            _ => "clear"
        };
    }

    private static string GetBorderStyle(BorderStyle style)
    {
        return style switch
        {
            BorderStyle.Single => "single",
            BorderStyle.Thick => "thick",
            BorderStyle.Double => "double",
            BorderStyle.Dotted => "dotted",
            BorderStyle.Dashed => "dash",
            BorderStyle.DotDash => "dotDash",
            BorderStyle.DotDotDash => "dotDotDash",
            BorderStyle.Triple => "triple",
            BorderStyle.ThinThickSmallGap => "thinThickSmallGap",
            BorderStyle.ThickThinSmallGap => "thickThinSmallGap",
            BorderStyle.ThinThickThinSmallGap => "thinThickThinSmallGap",
            BorderStyle.ThinThickMediumGap => "thinThickMediumGap",
            BorderStyle.ThickThinMediumGap => "thickThinMediumGap",
            BorderStyle.ThinThickThinMediumGap => "thinThickThinMediumGap",
            BorderStyle.ThinThickLargeGap => "thinThickLargeGap",
            BorderStyle.ThickThinLargeGap => "thickThinLargeGap",
            BorderStyle.ThinThickThinLargeGap => "thinThickThinLargeGap",
            BorderStyle.Wave => "wave",
            _ => "nil"
        };
    }

    private static bool IsLikelyMalformedBorder(BorderInfo border)
    {
        return border.Width > 96 && border.Color == 255;
    }

    private static int ConvertCharacterIndentToTwips(int characterUnits, int fontSizeHalfPoints)
    {
        int fontSizePoints = fontSizeHalfPoints > 0 ? Math.Max(1, fontSizeHalfPoints / 2) : 12;
        return (int)Math.Round(characterUnits * fontSizePoints * 10d);
    }

    private static string GetParagraphAlignmentValue(ParagraphAlignment alignment)
    {
        return alignment switch
        {
            ParagraphAlignment.Center => "center",
            ParagraphAlignment.Right => "right",
            ParagraphAlignment.Justify => "both",
            ParagraphAlignment.Distributed => "distribute",
            _ => "left"
        };
    }
}
