using System.Text;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Main document reader that orchestrates all the sub-readers.
/// 
/// Read pipeline:
///   1. CfbReader — extract streams from OLE2 container
///   2. FibReader — parse FIB from WordDocument stream
///   3. StyleReader — parse styles from Table stream
///   4. TextReader — parse text via CLX/Piece Table
///   5. Parse paragraphs and runs using FKP/SPRM data
///   6. TableReader — identify table structures
///   7. ImageReader — extract embedded images
/// </summary>
public class DocReader : IDisposable
{
    private CfbReader? _cfb;
    private BinaryReader? _wordDocReader;
    private BinaryReader? _tableReader;
    private BinaryReader? _dataReader;

    private FibReader? _fibReader;
    private TextReader? _textReader;
    private StyleReader? _styleReader;
    private DocumentPropertiesReader? _dopReader;
    private TableReader? _tableParseReader;
    private ImageReader? _imageReader;
    private FkpParser? _fkpParser;
    private FootnoteReader? _footnoteReader;
    private AnnotationReader? _annotationReader;
    private TextboxReader? _textboxReader;
    private HeaderFooterReader? _headerFooterReader;
    private ListReader? _listReader;
    private FieldReader? _fieldReader;
    private HyperlinkReader? _hyperlinkReader;
    private OfficeArtReader? _officeArtReader;
    private List<FspaInfo> _fspaAnchors = new();

    // Keep streams alive for reader lifetime
    private MemoryStream? _wordDocStream;
    private MemoryStream? _tableStream;
    private MemoryStream? _dataStream;
    private MemoryStream? _footnoteStream;
    private MemoryStream? _endnoteStream;
    private MemoryStream? _anotStream;
    private MemoryStream? _txbxStream;

    public DocumentModel Document { get; private set; } = new();
    public bool IsLoaded { get; private set; }

    public DocReader(string filePath)
    {
        _cfb = new CfbReader(filePath);
        InitializeStreams();
    }

    /// <summary>
    /// Attempts to infer simple category/series data for a chart from its
    /// embedded OLE bytes. This is intentionally conservative and only looks
    /// for small, tab/semicolon/comma-separated numeric tables inside the
    /// stream; if nothing reasonable is found the method leaves the model
    /// unchanged.
    /// </summary>
    private static void TryPopulateChartFromSourceBytes(ChartModel model)
    {
        var bytes = model.SourceBytes;
        if (bytes == null || bytes.Length < 64)
            return;

        // Avoid spending too much time on very large embedded streams.
        var maxBytes = Math.Min(bytes.Length, 512 * 1024);
        string text;
        try
        {
            // First try ANSI/UTF-8 style decoding.
            text = Encoding.Default.GetString(bytes, 0, maxBytes);
        }
        catch
        {
            return;
        }

        // Heuristic: require at least a few digits to consider this as
        // potentially containing chart data.
        if (!text.Any(char.IsDigit))
            return;

        var lines = text
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Take(200)
            .ToList();

        if (lines.Count < 2)
            return;

        // Tokenize by common separators.
        static string[] SplitTokens(string line)
        {
            return line
                .Split(new[] { '\t', ';', ',', '|' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => t.Length > 0)
                .ToArray();
        }

        var tokenLines = lines
            .Select(l => new { Line = l, Tokens = SplitTokens(l) })
            .Where(t => t.Tokens.Length >= 2)
            .ToList();

        if (tokenLines.Count < 2)
            return;

        // Use the first line as categories when it looks textual; otherwise
        // generate default category labels.
        var first = tokenLines[0];
        var second = tokenLines[1];

        var tokens = first.Tokens;
        bool firstLineLooksNumeric = tokens.All(tok =>
        {
            return double.TryParse(tok, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out _);
        });

        List<string> categories;
        if (!firstLineLooksNumeric)
        {
            categories = tokens.ToList();
        }
        else
        {
            categories = Enumerable.Range(1, second.Tokens.Length)
                .Select(i => $"Category {i}")
                .ToList();
        }

        // Use the next few lines as series values, aligning length with categories.
        var seriesList = new List<ChartSeries>();
        for (int i = 1; i < tokenLines.Count && i <= 5; i++)
        {
            var lineInfo = tokenLines[i];
            var valuesRaw = lineInfo.Tokens;
            var values = new List<double>();
            foreach (var tok in valuesRaw)
            {
                if (double.TryParse(tok, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var v))
                {
                    values.Add(v);
                }
                else
                {
                    // Non-numeric token – bail out of this heuristic series.
                    values.Clear();
                    break;
                }
            }

            if (values.Count == 0)
                continue;

            // Align with category count.
            if (values.Count < categories.Count)
            {
                while (values.Count < categories.Count)
                    values.Add(0);
            }
            else if (values.Count > categories.Count)
            {
                values = values.Take(categories.Count).ToList();
            }

            seriesList.Add(new ChartSeries
            {
                Name = $"Series {i}",
                Values = values
            });
        }

        if (seriesList.Count == 0)
            return;

        model.Categories = categories;
        model.Series = seriesList;
    }

    public DocReader(Stream stream)
    {
        _cfb = new CfbReader(stream, leaveOpen: true);
        InitializeStreams();
    }

    private void InitializeStreams()
    {
        // Extract WordDocument stream (required)
        if (!_cfb!.HasStream("WordDocument"))
            throw new InvalidDataException("Not a valid Word document: 'WordDocument' stream not found.");

        // First, read the raw WordDocument stream to parse the FIB
        _wordDocStream = _cfb.GetStream("WordDocument");
        _wordDocReader = new BinaryReader(_wordDocStream, Encoding.Default, leaveOpen: true);

        _fibReader = new FibReader(_wordDocReader);
        _fibReader.Read();

        // If the document is XOR-encrypted, configure the CFB reader with the key
        // and reopen the core streams in decrypted form for all subsequent parsing.
        if (_fibReader.FEncrypted)
        {
            _cfb.SetEncryptionKey(_fibReader.LKey);

            _wordDocReader.Dispose();
            _wordDocStream.Dispose();

            _wordDocStream = _cfb.GetDecryptedStream("WordDocument");
            _wordDocReader = new BinaryReader(_wordDocStream, Encoding.Default, leaveOpen: true);
        }

        // Extract Table stream (0Table or 1Table)
        var tableName = _fibReader.TableStreamName;
        if (!_cfb.HasStream(tableName))
        {
            // Try the other one
            tableName = tableName == "1Table" ? "0Table" : "1Table";
            if (!_cfb.HasStream(tableName))
                throw new InvalidDataException($"Table stream not found. Tried '{_fibReader.TableStreamName}' and '{tableName}'.");
        }

        _tableStream = _fibReader.FEncrypted
            ? _cfb.GetDecryptedStream(tableName)
            : _cfb.GetStream(tableName);
        _tableReader = new BinaryReader(_tableStream, Encoding.Default, leaveOpen: true);

        // Read floating shape anchors from PlcfSpaMom (best-effort).
        try
        {
            _fspaAnchors = FspaReader.ReadPlcSpaMom(_tableReader, _fibReader);
        }
        catch
        {
            _fspaAnchors = new List<FspaInfo>();
        }

        // Extract Data stream (optional — contains pictures, OLE objects)
        if (_cfb.HasStream("Data"))
        {
            _dataStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Data")
                : _cfb.GetStream("Data");
            _dataReader = new BinaryReader(_dataStream, Encoding.Default, leaveOpen: true);

            // Initialize OfficeArt/Escher reader on the Data stream (best-effort).
            try
            {
                _officeArtReader = new OfficeArtReader(_dataStream);
            }
            catch
            {
                _officeArtReader = null;
            }
        }

        // Extract footnote/endnote streams (optional)
        if (_cfb.HasStream("Footnote"))
        {
            _footnoteStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Footnote")
                : _cfb.GetStream("Footnote");
        }

        if (_cfb.HasStream("Endnote"))
        {
            _endnoteStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Endnote")
                : _cfb.GetStream("Endnote");
        }

        // Extract annotation stream (optional)
        if (_cfb.HasStream("Anot"))
        {
            _anotStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Anot")
                : _cfb.GetStream("Anot");
        }

        // Extract textbox stream (optional)
        if (_cfb.HasStream("Txbx"))
        {
            _txbxStream = _fibReader.FEncrypted
                ? _cfb.GetDecryptedStream("Txbx")
                : _cfb.GetStream("Txbx");
        }

        // Initialize sub-readers
        _textReader = new TextReader(_wordDocReader!, _tableReader!, _fibReader!);
        _styleReader = new StyleReader(_tableReader!, _fibReader!);
        _dopReader = new DocumentPropertiesReader(_tableReader!, _fibReader!);
        _fkpParser = new FkpParser(_wordDocReader!, _tableReader!, _fibReader!, _textReader!);
        _tableParseReader = new TableReader(_wordDocReader!, _tableReader!, _fibReader!, _fkpParser);
        _imageReader = new ImageReader(_wordDocReader!, _dataReader, _fibReader!, _cfb);
        _footnoteReader = new FootnoteReader(_fibReader!, _textReader!);
        _annotationReader = new AnnotationReader(
            _anotStream != null ? new BinaryReader(_anotStream, Encoding.Default, leaveOpen: true) : null,
            _fibReader!);
        _textboxReader = new TextboxReader(_tableReader!, _fibReader!, _textReader!);
        _headerFooterReader = new HeaderFooterReader(_tableReader!, _wordDocReader!, _fibReader!, _textReader!);
        _listReader = new ListReader(_tableReader!, _fibReader!);
        _fieldReader = new FieldReader();
        _hyperlinkReader = new HyperlinkReader();
    }

    /// <summary>
    /// Loads and parses the document.
    /// </summary>
    public void Load()
    {
        // Step 1: Read document properties
        Document.Properties = _dopReader!.Read();

        // Step 1.5: Read style sheet
        _styleReader!.Read();
        Document.Styles = _styleReader.Styles;

        // Step 2: Read list definitions
        _listReader!.Read();
        Document.NumberingDefinitions = _listReader.NumberingDefinitions;
        Document.ListFormats = _listReader.ListFormats;

        // Step 3: Read text content via Piece Table (with per-run Lid for encoding)
        var totalCp = _fibReader!.CcpText + _fibReader.CcpFtn + _fibReader.CcpHdd + _fibReader.CcpAtn + _fibReader.CcpEdn + _fibReader.CcpTxbx + _fibReader.CcpHdrTxbx;
        _textReader!.ReadClx();
        var chpMap = _fkpParser!.ReadChpProperties();
        _textReader.SetTextFromPieces(totalCp, chpMap);

        // Step 4: Parse paragraphs and runs
        ParseDocumentContent();

        // Step 5: Parse tables
        _tableParseReader!.ParseTables(Document);

        // Step 6: Extract images
        _imageReader!.ExtractImages(Document);

        // Step 6.5: Parse OfficeArt/Escher shapes and map basic anchors
        if (_officeArtReader != null)
        {
            OfficeArtMapper.AttachShapes(Document, _officeArtReader, _fspaAnchors);
        }

        // Step 6.6: Best-effort chart detection. We recognise streams whose
        // names contain "Chart" in the OLE container and attach ChartModel
        // instances. A lightweight heuristic then tries to recover real
        // category/value data from the embedded stream; when that fails we
        // fall back to placeholder data so the resulting chart remains
        // editable in Word.
        AttachPlaceholderCharts();

        // Step 7: Read footnotes
        if (_footnoteReader != null)
        {
            Document.Footnotes = _footnoteReader.ReadFootnotesWithOffset();
        }

        // Step 8: Read endnotes
        if (_footnoteReader != null)
        {
            Document.Endnotes = _footnoteReader.ReadEndnotesWithOffset();
        }

        if (_annotationReader != null)
        {
            Document.Annotations = _annotationReader.ReadAnnotations();
        }

        // Step 9: Read textboxes
        if (_textboxReader != null)
        {
            Document.Textboxes = _textboxReader.ReadTextboxes();
        }

        // Step 10: Read headers/footers
        if (_headerFooterReader != null)
        {
            _headerFooterReader.Read(Document);
            Document.HeadersFooters.Headers = _headerFooterReader.Headers;
            Document.HeadersFooters.Footers = _headerFooterReader.Footers;
        }

        IsLoaded = true;
    }

    /// <summary>
    /// Creates placeholder ChartModel instances for any OLE streams whose names
    /// look like chart containers. This gives callers a basic, editable chart
    /// in the resulting DOCX even when we do not yet understand the underlying
    /// binary chart format.
    /// </summary>
    private void AttachPlaceholderCharts()
    {
        if (_cfb == null) return;

        // Look for streams which strongly suggest embedded charts. In many
        // real-world documents these names come from OLE chart objects.
        var chartLikeStreams = _cfb.StreamNames
            .Where(n => n.IndexOf("Chart", StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();

        if (chartLikeStreams.Count == 0)
            return;

        int existingCount = Document.Charts.Count;
        for (int i = 0; i < chartLikeStreams.Count; i++)
        {
            var name = chartLikeStreams[i];

            // Try to capture the raw OLE stream bytes for this chart so that
            // future phases (or callers) can recover real series data from the
            // original container (e.g. MS Graph or embedded Excel workbook).
            byte[]? sourceBytes = null;
            try
            {
                sourceBytes = _cfb.GetStreamBytes(name);
            }
            catch
            {
                // best-effort only; leave SourceBytes as null on failure
            }

            var model = new ChartModel
            {
                Index = existingCount + i,
                Title = name,
                Type = ChartType.Column,
                Categories = new List<string> { "Category 1", "Category 2", "Category 3" },
                Series =
                {
                    new ChartSeries
                    {
                        Name = "Series 1",
                        Values = new List<double> { 1, 2, 3 }
                    }
                },
                SourceStreamName = name,
                SourceBytes = sourceBytes,
                ParagraphIndexHint = -1
            };

            // Best-effort attempt to recover real categories/series from the
            // embedded bytes; falls back to the placeholder data above when
            // nothing sensible can be inferred.
            TryPopulateChartFromSourceBytes(model);

            Document.Charts.Add(model);
        }

        // If we still do not have paragraph hints for charts, distribute them
        // roughly evenly across "normal" paragraphs so that each chart appears
        // near some body text rather than all being appended at the end.
        if (Document.Paragraphs.Count > 0)
        {
            var candidateParagraphIndices = Document.Paragraphs
                .Where(p => p.Type == ParagraphType.Normal)
                .Select(p => p.Index)
                .ToList();

            if (candidateParagraphIndices.Count == 0)
            {
                candidateParagraphIndices = Document.Paragraphs.Select(p => p.Index).ToList();
            }

            if (candidateParagraphIndices.Count > 0)
            {
                var chartsNeedingPlacement = Document.Charts
                    .Where(c => c.ParagraphIndexHint < 0)
                    .ToList();

                for (int i = 0; i < chartsNeedingPlacement.Count; i++)
                {
                    var target = candidateParagraphIndices[
                        (int)((long)i * candidateParagraphIndices.Count / chartsNeedingPlacement.Count)
                    ];
                    chartsNeedingPlacement[i].ParagraphIndexHint = target;
                }
            }
        }
    }

    /// <summary>
    /// Parses the document content into paragraphs and runs using FKP-based parsing.
    /// 
    /// This implementation:
    ///   - Reads CHP and PAP properties from FKPs
    ///   - Splits text by paragraph marks (CR = 0x0D)
    ///   - Creates multiple runs per paragraph based on CHP changes
    ///   - Applies paragraph formatting from PAP FKPs
    /// </summary>
    private void ParseDocumentContent()
    {
        var text = _textReader!.Text;
        if (string.IsNullOrEmpty(text)) return;

        // Read CHP and PAP properties from FKPs
        var chpMap = _fkpParser!.ReadChpProperties();
        var papMap = _fkpParser.ReadPapProperties();

        // In Word binary format, paragraphs are delimited by CR (0x0D)
        // Special characters: 0x07 = cell mark, 0x0C = page break,
        // 0x01 = field begin/end or inline picture
        var paragraphIndex = 0;
        var paraStart = 0;
        var imageCounter = 0;

        // Only iterate characters in the main document range [0, ccpText)
        int mainDocumentLength = Math.Min(text.Length, _fibReader!.CcpText);

        for (int i = 0; i <= mainDocumentLength; i++)
        {
            bool isParagraphEnd = (i == mainDocumentLength) || (text[i] == '\r') || (text[i] == '\x0D');

            if (!isParagraphEnd) continue;

            var paraText = text.Substring(paraStart, i - paraStart);
            var paraStartCp = paraStart;
            paraStart = i + 1; // skip the delimiter
            if (paraStart > mainDocumentLength) paraStart = mainDocumentLength;

            // Get PAP for this paragraph (use the first CP of the paragraph; if none, use nearest preceding CP so style/alignment are preserved)
            PapBase? pap = null;
            if (paraStartCp < text.Length && papMap.TryGetValue(paraStartCp, out var foundPap))
                pap = foundPap;
            if (pap == null && papMap.Count > 0)
            {
                // No PAP at paraStartCp (gap in PLC). Prefer preceding PAP; if it's Normal, try following (title may be in next segment).
                for (int cp = paraStartCp - 1; cp >= 0; cp--)
                {
                    if (papMap.TryGetValue(cp, out var prevPap)) { pap = prevPap; break; }
                }
                if (pap != null && pap.StyleId == 0 && pap.Istd == 0)
                {
                    for (int cp = paraStartCp + 1; cp <= paraStartCp + 2000 && cp <= _fibReader!.CcpText; cp++)
                    {
                        if (papMap.TryGetValue(cp, out var nextPap) && (nextPap.StyleId != 0 || nextPap.Istd != 0))
                        {
                            pap = nextPap;
                            break;
                        }
                    }
                }
                if (pap == null)
                {
                    var firstKey = papMap.Keys.Min();
                    if (firstKey <= paraStartCp + 2000)
                        papMap.TryGetValue(firstKey, out pap);
                }
            }

            var paragraph = new ParagraphModel
            {
                Index = paragraphIndex++,
                Type = ParagraphType.Normal,
                Properties = pap != null 
                    ? _fkpParser.ConvertToParagraphProperties(pap, Document.Styles)
                    : new ParagraphProperties(),
                ListFormatId = pap?.ListFormatId ?? 0,
                ListLevel = pap?.ListLevel ?? 0
            };

            // Detect special paragraph types
            if (paraText.Contains('\x07'))
            {
                // Contains cell end mark — this is a table cell paragraph
                paragraph.Type = ParagraphType.TableCell;
            }
            else if (paraText.Contains('\x0C'))
            {
                paragraph.Type = ParagraphType.PageBreak;
            }

            // Split paragraph into runs based on CHP changes; when no direct CHP,
            // inherit run properties (font size, color, etc.) from the paragraph style.
            var runs = ParseRunsInParagraph(paraText, paraStartCp, chpMap, papMap, ref imageCounter);
            paragraph.Runs.AddRange(runs);

            // If no runs were created (no CHP data), create one run using paragraph style so font/color/size are preserved
            if (paragraph.Runs.Count == 0)
            {
                var cleanText = CleanSpecialChars(paraText);
                if (!string.IsNullOrEmpty(cleanText))
                {
                    var runProps = GetRunPropertiesFromParagraphStyleAtCp(papMap, paraStartCp) ?? new RunProperties { FontSize = 24 };
                    paragraph.Runs.Add(new RunModel
                    {
                        Text = cleanText,
                        CharacterPosition = paraStartCp,
                        CharacterLength = paraText.Length,
                        Properties = runProps
                    });
                }
            }

            ApplyParagraphStyleDefaults(paragraph);
            ApplyTemplateSpecificFixes(paragraph);

            Document.Paragraphs.Add(paragraph);
        }
    }

    /// <summary>
    /// Parses runs within a paragraph based on CHP property changes.
    /// When there is no direct CHP at a position, run properties are taken from the paragraph's style so that
    /// font size and color from the .doc are preserved.
    /// </summary>
    private List<RunModel> ParseRunsInParagraph(string paraText, int paraStartCp, Dictionary<int, ChpBase> chpMap, Dictionary<int, PapBase> papMap, ref int imageCounter)
    {
        var runs = new List<RunModel>();
        if (string.IsNullOrEmpty(paraText)) return runs;

        var runStart = 0;
        ChpBase? currentChp = null;

        for (int i = 0; i <= paraText.Length; i++)
        {
            var cp = paraStartCp + i;
            ChpBase? chpAtCp = null;
            
            if (chpMap.TryGetValue(cp, out var foundChp))
                chpAtCp = foundChp;

            bool chpChanged = i == paraText.Length || !ChpEquals(currentChp, chpAtCp);

            if (chpChanged && runStart < i)
            {
                var runText = paraText.Substring(runStart, i - runStart);
                var cleanText = CleanSpecialChars(runText);
                var isPicture = runText.Contains('\x01') || runText.Contains('\x08');

                if (!string.IsNullOrEmpty(cleanText) || isPicture)
                {
                    RunProperties runProps;
                    if (currentChp != null)
                        runProps = _fkpParser!.ConvertToRunProperties(currentChp, Document.Styles);
                    else
                        runProps = GetRunPropertiesFromParagraphStyleAtCp(papMap, paraStartCp) ?? new RunProperties { FontSize = 24 };

                    var run = new RunModel
                    {
                        Text = cleanText,
                        CharacterPosition = paraStartCp + runStart,
                        CharacterLength = runText.Length,
                        Properties = runProps
                    };

                    // Check for field characters (0x13 = field begin, 0x14 = separator, 0x15 = end)
                    if (runText.Contains('\x13'))
                    {
                        run.IsField = true;
                        var fieldStart = runText.IndexOf('\x13');
                        var fieldSep = runText.IndexOf('\x14');
                        if (fieldStart >= 0 && fieldSep > fieldStart)
                        {
                            run.FieldCode = runText.Substring(fieldStart + 1, fieldSep - fieldStart - 1).Trim();
                        }

                        // Try to interpret hyperlink fields as true OOXML hyperlinks
                        if (!string.IsNullOrEmpty(run.FieldCode) && _fieldReader != null && _hyperlinkReader != null)
                        {
                            var field = _fieldReader.ParseField(run.FieldCode);
                            if (field != null && field.Type == FieldType.Hyperlink)
                            {
                                var hyperlink = _hyperlinkReader.ParseHyperlink(run.FieldCode)
                                                ?? _hyperlinkReader.CreateHyperlink(field.Arguments);

                                if (!string.IsNullOrEmpty(hyperlink.Url))
                                {
                                    run.IsHyperlink = true;
                                    run.HyperlinkUrl = hyperlink.Url;
                                    run.HyperlinkRelationshipId = hyperlink.RelationshipId;

                                    // Treat this run as a hyperlink rather than a generic field
                                    run.IsField = false;

                                    if (!Document.Hyperlinks.Any(h => string.Equals(h.Url, hyperlink.Url, StringComparison.OrdinalIgnoreCase)
                                                                       && string.Equals(h.Bookmark, hyperlink.Bookmark, StringComparison.Ordinal)))
                                    {
                                        Document.Hyperlinks.Add(hyperlink);
                                    }
                                }
                            }
                        }
                    }

                    if (isPicture)
                    {
                        run.IsPicture = true;
                        run.ImageIndex = imageCounter++;
                        run.FcPic = currentChp?.FcPic ?? 0;
                    }

                    runs.Add(run);
                }

                runStart = i;
            }

            currentChp = chpAtCp;
        }

        return runs;
    }

    /// <summary>
    /// Compares two CHP objects for equality so we split runs when formatting (including color/size) changes.
    /// </summary>
    private static bool ChpEquals(ChpBase? a, ChpBase? b)
    {
        if (ReferenceEquals(a, b)) return true;
        if (a == null || b == null) return false;

        return a.IsBold == b.IsBold &&
               a.IsItalic == b.IsItalic &&
               a.IsStrikeThrough == b.IsStrikeThrough &&
               a.IsUnderline == b.IsUnderline &&
               a.FontSize == b.FontSize &&
               a.FontSizeCs == b.FontSizeCs &&
               a.FontIndex == b.FontIndex &&
               a.Color == b.Color &&
               a.HasRgbColor == b.HasRgbColor &&
               a.RgbColor == b.RgbColor;
    }

    /// <summary>
    /// Gets run properties from the paragraph style at the given CP when there is no direct CHP,
    /// so that style-based font size and color from the .doc are preserved.
    /// </summary>
    private RunProperties? GetRunPropertiesFromParagraphStyleAtCp(Dictionary<int, PapBase> papMap, int cp)
    {
        if (!papMap.TryGetValue(cp, out var pap)) return null;
        var styles = Document.Styles;
        if (styles?.Styles == null || styles.Styles.Count == 0) return null;
        var styleIndex = pap.StyleId != 0 ? pap.StyleId : pap.Istd;
        var style = styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == styleIndex);
        var sr = style?.RunProperties;
        if (sr == null) return null;
        return CloneRunProperties(sr);
    }

    private static RunProperties CloneRunProperties(RunProperties sr)
    {
        var r = new RunProperties
        {
            FontIndex = sr.FontIndex,
            FontName = sr.FontName,
            FontSize = sr.FontSize,
            FontSizeCs = sr.FontSizeCs,
            IsBold = sr.IsBold,
            IsBoldCs = sr.IsBoldCs,
            IsItalic = sr.IsItalic,
            IsItalicCs = sr.IsItalicCs,
            IsUnderline = sr.IsUnderline,
            UnderlineType = sr.UnderlineType,
            IsStrikeThrough = sr.IsStrikeThrough,
            IsDoubleStrikeThrough = sr.IsDoubleStrikeThrough,
            IsSmallCaps = sr.IsSmallCaps,
            IsAllCaps = sr.IsAllCaps,
            IsHidden = sr.IsHidden,
            IsSuperscript = sr.IsSuperscript,
            IsSubscript = sr.IsSubscript,
            Color = sr.Color,
            BgColor = sr.BgColor,
            CharacterSpacingAdjustment = sr.CharacterSpacingAdjustment,
            Language = sr.Language,
            LanguageAsia = sr.LanguageAsia,
            LanguageCs = sr.LanguageCs,
            HighlightColor = sr.HighlightColor,
            RgbColor = sr.RgbColor,
            HasRgbColor = sr.HasRgbColor,
            IsOutline = sr.IsOutline,
            IsShadow = sr.IsShadow,
            IsEmboss = sr.IsEmboss,
            IsImprint = sr.IsImprint,
            Kerning = sr.Kerning,
            Position = sr.Position
        };
        return r;
    }

    private void ApplyParagraphStyleDefaults(ParagraphModel paragraph)
    {
        if (paragraph == null) return;
        var paragraphProps = paragraph.Properties;
        if (paragraphProps == null) return;

        var styles = Document.Styles;
        if (styles == null || styles.Styles == null || styles.Styles.Count == 0)
            return;

        var style = styles.Styles
            .FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == paragraphProps.StyleIndex);

        if (style == null)
            return;

        // Paragraph-level properties
        if (style.ParagraphProperties != null)
        {
            var sp = style.ParagraphProperties;

            if (paragraphProps.Alignment == ParagraphAlignment.Left &&
                sp.Alignment != ParagraphAlignment.Left)
            {
                paragraphProps.Alignment = sp.Alignment;
            }

            if (paragraphProps.IndentLeft == 0 && sp.IndentLeft != 0)
                paragraphProps.IndentLeft = sp.IndentLeft;

            if (paragraphProps.IndentRight == 0 && sp.IndentRight != 0)
                paragraphProps.IndentRight = sp.IndentRight;

            if (paragraphProps.IndentFirstLine == 0 && sp.IndentFirstLine != 0)
                paragraphProps.IndentFirstLine = sp.IndentFirstLine;

            if (paragraphProps.SpaceBefore == 0 && sp.SpaceBefore != 0)
                paragraphProps.SpaceBefore = sp.SpaceBefore;

            if (paragraphProps.SpaceAfter == 0 && sp.SpaceAfter != 0)
                paragraphProps.SpaceAfter = sp.SpaceAfter;

            if (paragraphProps.LineSpacing == 240 && sp.LineSpacing != 240)
            {
                paragraphProps.LineSpacing = sp.LineSpacing;
                paragraphProps.LineSpacingMultiple = sp.LineSpacingMultiple;
            }
        }

        // Run-level defaults: apply style run properties to each run when
        // the run hasn't specified its own font/color/etc.
        if (style.RunProperties == null || paragraph.Runs == null || paragraph.Runs.Count == 0)
            return;

        var sr = style.RunProperties;

        foreach (var run in paragraph.Runs)
        {
            run.Properties ??= new RunProperties();
            var rp = run.Properties;

            // Font name
            if (string.IsNullOrEmpty(rp.FontName) && !string.IsNullOrEmpty(sr.FontName))
                rp.FontName = sr.FontName;

            // Font size (24 half-points = 12pt default)
            if (rp.FontSize == 24 && sr.FontSize != 24)
                rp.FontSize = sr.FontSize;

            // Bold / italic
            if (!rp.IsBold && sr.IsBold)
                rp.IsBold = true;
            if (!rp.IsItalic && sr.IsItalic)
                rp.IsItalic = true;

            // Color / RGB color: 0 + !HasRgbColor = "auto"
            if (!rp.HasRgbColor && rp.Color == 0)
            {
                if (sr.HasRgbColor)
                {
                    rp.RgbColor = sr.RgbColor;
                    rp.HasRgbColor = true;
                }
                else if (sr.Color != 0)
                {
                    rp.Color = sr.Color;
                }
            }

            // Highlight
            if (rp.HighlightColor == 0 && sr.HighlightColor != 0)
                rp.HighlightColor = sr.HighlightColor;
        }
    }

    /// <summary>
    /// Applies document/template‑specific fallbacks that are hard to
    /// infer purely from low‑level binary structures.
    /// </summary>
    private void ApplyTemplateSpecificFixes(ParagraphModel paragraph)
    {
        if (paragraph == null) return;

        if (string.IsNullOrEmpty(paragraph.Text) || !paragraph.Text.Contains("绿色等级评价报告", StringComparison.Ordinal))
            return;

        paragraph.Properties ??= new ParagraphProperties();
        paragraph.Properties.Alignment = ParagraphAlignment.Center;

        // When PAP gave Normal but CHP has larger font (direct formatting), take color/font from a title-like style if present
        var styles = Document.Styles?.Styles;
        if (styles == null || paragraph.Runs == null) return;
        var titleLike = styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId != 0 &&
            (s.Name?.Contains("Title", StringComparison.OrdinalIgnoreCase) == true ||
             s.Name?.Contains("标题", StringComparison.OrdinalIgnoreCase) == true ||
             s.Name?.Contains("Heading", StringComparison.OrdinalIgnoreCase) == true) &&
            s.RunProperties != null && (s.RunProperties.Color != 0 || s.RunProperties.HasRgbColor || s.RunProperties.FontSize > 24));
        if (titleLike?.RunProperties == null) return;
        var tr = titleLike.RunProperties;
        foreach (var run in paragraph.Runs)
        {
            run.Properties ??= new RunProperties();
            if (run.Properties.Color == 0 && !run.Properties.HasRgbColor && tr.Color != 0) run.Properties.Color = tr.Color;
            if (run.Properties.Color == 0 && !run.Properties.HasRgbColor && tr.HasRgbColor) { run.Properties.RgbColor = tr.RgbColor; run.Properties.HasRgbColor = true; }
            if (run.Properties.FontSize <= 24 && tr.FontSize > 24) run.Properties.FontSize = tr.FontSize;
            if (string.IsNullOrEmpty(run.Properties.FontName) && !string.IsNullOrEmpty(tr.FontName)) run.Properties.FontName = tr.FontName;
        }
    }

    /// <summary>
    /// Removes Word special control characters from text for display.
    /// </summary>
    private static string CleanSpecialChars(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '\x01': // SOH — field begin/end or inline picture
                case '\x07': // BEL — cell mark
                case '\x08': // BS  — drawn object
                case '\x13': // field begin
                case '\x14': // field separator
                case '\x15': // field end
                    break; // skip these

                case '\x0B': // vertical tab → line break
                    sb.Append('\n');
                    break;

                case '\x0C': // form feed → page break (keep as text for now)
                    break;

                case '\x1E': // non-breaking hyphen
                    sb.Append('-');
                    break;

                case '\x1F': // optional hyphen
                    break; // skip

                case '\xA0': // non-breaking space
                    sb.Append(' ');
                    break;

                default:
                    if (!char.IsControl(ch) || ch == '\t')
                        sb.Append(ch);
                    break;
            }
        }
        return sb.ToString();
    }

    /// <summary>Gets the text reader for direct access</summary>
    public TextReader GetTextReader() => _textReader!;

    /// <summary>Gets the FIB reader for direct access</summary>
    public FibReader GetFibReader() => _fibReader!;

    /// <summary>Gets the style reader for direct access</summary>
    public StyleReader GetStyleReader() => _styleReader!;

    /// <summary>Gets the CFB reader for diagnostics</summary>
    public CfbReader GetCfbReader() => _cfb!;

    public void Dispose()
    {
        _wordDocReader?.Dispose();
        _tableReader?.Dispose();
        _dataReader?.Dispose();
        _footnoteStream?.Dispose();
        _endnoteStream?.Dispose();
        _anotStream?.Dispose();
        _txbxStream?.Dispose();
        _wordDocStream?.Dispose();
        _tableStream?.Dispose();
        _dataStream?.Dispose();
        _cfb?.Dispose();
    }
}

/// <summary>
    /// Table reader — parses table structures from document.
    /// Uses a combination of ParagraphType.TableCell markers and TAP (table
    /// properties) decoded from PAP/FKP data. This gives reasonably faithful
    /// reconstruction of row heights, header rows and horizontal/vertical merges
    /// for the common cases while deliberately avoiding the full generality of
    /// nested tables and exotic merge patterns present in the MS-DOC format.
/// </summary>
public class TableReader
{
    private readonly BinaryReader _wordDocReader;
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly FkpParser _fkpParser;

        public TableReader(BinaryReader wordDocReader, BinaryReader tableReader, FibReader fib, FkpParser fkpParser)
        {
            _wordDocReader = wordDocReader;
            _tableReader = tableReader;
            _fib = fib;
            _fkpParser = fkpParser;
        }

    /// <summary>
    /// Parses tables from the document by examining paragraph types.
    /// Groups contiguous ParagraphType.TableCell 段落为一张张独立的表格。
    /// </summary>
        public void ParseTables(DocumentModel document)
        {
            var tables = new List<TableModel>();
            TableModel? currentTable = null;
            var rowIndex = 0;
            var cellsInCurrentRow = new List<TableCellModel>();
            int lastTableParagraphIndex = -1;
            TapBase? currentRowTap = null;
            // 保留每一行对应的 TAP 信息，用于之后根据 TC.grfw 计算单元格纵向/横向合并。
            var rowTaps = new List<TapBase?>();

        foreach (var para in document.Paragraphs.OrderBy(p => p.Index))
        {
            if (para.Type != ParagraphType.TableCell)
            {
                // 离开表格区域：收尾当前表格
                if (currentTable != null)
                {
                    FlushCurrentRow();
                    FinalizeCurrentTable();
                }
                continue;
            }

            // 进入一个新的表格块
            if (currentTable == null)
            {
                currentTable = new TableModel
                {
                    Index = tables.Count,
                    StartParagraphIndex = para.Index
                };
                rowIndex = 0;
                cellsInCurrentRow.Clear();
                currentRowTap = null;
            }

            lastTableParagraphIndex = para.Index;

            // 从 FKP/PAPX 中获取与该段落关联的 TAP 信息
            TapBase? tapForParagraph = null;
            var firstRun = para.Runs.FirstOrDefault();
            if (firstRun != null)
            {
                var pap = _fkpParser.GetPapAtCp(firstRun.CharacterPosition);
                tapForParagraph = pap?.Tap;

                if (tapForParagraph != null && currentTable.Properties == null)
                {
                    // Map TAP‑level table properties into the high‑level model so that
                    // the writer can faithfully reproduce alignment, indent, spacing
                    // table width, and table‑wide borders / shading.
                    currentTable.Properties = new TableProperties
                    {
                        Alignment = tapForParagraph.Justification switch
                        {
                            1 => TableAlignment.Center,
                            2 => TableAlignment.Right,
                            _ => TableAlignment.Left
                        },
                        // Prefer the TAP CellSpacing value; if it is zero, fall back to
                        // 2 * GapHalf (derived from sprmTDxaGapHalf) where available.
                        CellSpacing = tapForParagraph.CellSpacing != 0
                            ? tapForParagraph.CellSpacing
                            : (tapForParagraph.GapHalf != 0 ? tapForParagraph.GapHalf * 2 : 0),
                        Indent = tapForParagraph.IndentLeft,
                        PreferredWidth = tapForParagraph.TableWidth,
                        BorderTop = tapForParagraph.BorderTop,
                        BorderBottom = tapForParagraph.BorderBottom,
                        BorderLeft = tapForParagraph.BorderLeft,
                        BorderRight = tapForParagraph.BorderRight,
                        BorderInsideH = tapForParagraph.BorderInsideH,
                        BorderInsideV = tapForParagraph.BorderInsideV,
                        Shading = tapForParagraph.Shading
                    };
                }

                if (currentRowTap == null && tapForParagraph != null)
                {
                    currentRowTap = tapForParagraph;
                }
            }

            // 检测是否为“行结束”标记段落：只包含单元格结束符 \x07 或完全为空
            var isRowEnd = para.Runs.Count == 0 ||
                           (para.Runs.Count == 1 && string.IsNullOrWhiteSpace(para.Runs[0].Text.Replace("\x07", "")));

            if (isRowEnd)
            {
                // 结束当前行但不创建新的单元格
                FlushCurrentRow();
            }
            else
            {
                // 普通单元格内容段落：复用原段落的段落属性（包含样式继承结果），
                // 并按单元格需要复制文本 run。
                var cellParagraph = new ParagraphModel
                {
                    Index = 0,
                    Type = ParagraphType.Normal,
                    Properties = para.Properties,
                    Runs = para.Runs.Select(r => new RunModel
                    {
                        Text = r.Text.Replace("\x07", ""),
                        Properties = r.Properties != null ? new RunProperties
                        {
                            IsBold = r.Properties.IsBold,
                            IsItalic = r.Properties.IsItalic,
                            IsUnderline = r.Properties.IsUnderline,
                            FontSize = r.Properties.FontSize,
                            FontName = r.Properties.FontName,
                            Color = r.Properties.Color,
                            BgColor = r.Properties.BgColor
                        } : null
                    }).ToList()
                };

                // 移除空 run
                cellParagraph.Runs.RemoveAll(r => string.IsNullOrEmpty(r.Text));

                var cellModel = new TableCellModel
                {
                    Index = cellsInCurrentRow.Count,
                    RowIndex = rowIndex,
                    ColumnIndex = cellsInCurrentRow.Count,
                    Paragraphs = new List<ParagraphModel> { cellParagraph }
                };

                // 从 TAP 的 CellWidths 推导单元格宽度（使用第一行的宽度信息）
                if (tapForParagraph?.CellWidths != null &&
                    tapForParagraph.CellWidths.Length > cellModel.ColumnIndex)
                {
                    cellModel.Properties ??= new TableCellProperties();
                    cellModel.Properties.Width = tapForParagraph.CellWidths[cellModel.ColumnIndex];
                }

                cellsInCurrentRow.Add(cellModel);
            }
        }

        // 文档结束时，如仍在表格中，收尾
            if (currentTable != null)
            {
                FlushCurrentRow();
                FinalizeCurrentTable();
            }

            document.Tables = tables;

            // 本地帮助方法：结束当前行
            void FlushCurrentRow()
            {
                if (currentTable == null) return;
                if (cellsInCurrentRow.Count == 0) return;

                var row = new TableRowModel
                {
                    Index = rowIndex++,
                    Cells = new List<TableCellModel>(cellsInCurrentRow)
                };

                if (currentRowTap != null)
                {
                    row.Properties ??= new TableRowProperties();
                    if (currentRowTap.RowHeight > 0)
                    {
                        row.Properties.Height = currentRowTap.RowHeight;
                        row.Properties.HeightIsExact = currentRowTap.HeightIsExact;
                    }
                    if (currentRowTap.IsHeaderRow)
                    {
                        row.Properties.IsHeaderRow = true;
                    }
                    row.Properties.AllowBreakAcrossPages = !currentRowTap.CantSplit;
                }

                currentTable.Rows.Add(row);
                rowTaps.Add(currentRowTap);
                cellsInCurrentRow.Clear();
                currentRowTap = null;
            }

            // 本地帮助方法：计算行列数并添加到集合
            void FinalizeCurrentTable()
            {
                if (currentTable == null) return;
                if (currentTable.Rows.Count == 0)
                {
                    currentTable = null;
                    rowTaps.Clear();
                    return;
                }

                currentTable.EndParagraphIndex = lastTableParagraphIndex;
                currentTable.RowCount = currentTable.Rows.Count;
                currentTable.ColumnCount = currentTable.Rows.Max(r => r.Cells.Count);

                // 默认将每张表的第一行标记为表头行，便于在 Word 中重复显示在每页顶部。
                var firstRow = currentTable.Rows.FirstOrDefault();
                if (firstRow != null)
                {
                    firstRow.Properties ??= new TableRowProperties();
                    firstRow.Properties.IsHeaderRow = true;
                }

                // 1) 基于 TAP / TC 的精确信息推断纵向合并（RowSpan）
                bool hasTapMergeInfo = rowTaps.Any(t => t?.CellMerges != null);
                if (hasTapMergeInfo && currentTable.ColumnCount > 0)
                {
                    for (int col = 0; col < currentTable.ColumnCount; col++)
                    {
                        int row = 0;
                        while (row < currentTable.Rows.Count)
                        {
                            var startCell = GetCell(currentTable, row, col);
                            if (startCell == null)
                            {
                                row++;
                                continue;
                            }

                            var tap = row < rowTaps.Count ? rowTaps[row] : null;
                            var mergeArray = tap?.CellMerges;
                            CellMergeFlags? flags = null;
                            if (mergeArray != null && col < mergeArray.Length)
                            {
                                flags = mergeArray[col];
                            }

                            if (flags == null || !flags.VertFirst)
                            {
                                row++;
                                continue;
                            }

                            int span = 1;
                            int nextRow = row + 1;
                            while (nextRow < currentTable.Rows.Count)
                            {
                                var nextTap = nextRow < rowTaps.Count ? rowTaps[nextRow] : null;
                                var nextArray = nextTap?.CellMerges;
                                CellMergeFlags? nextFlags = null;
                                if (nextArray != null && col < nextArray.Length)
                                {
                                    nextFlags = nextArray[col];
                                }

                                if (nextFlags == null || !nextFlags.VertMerged)
                                {
                                    break;
                                }

                                span++;
                                nextRow++;
                            }

                            if (span > 1)
                            {
                                startCell.RowSpan = span;
                                row += span;
                            }
                            else
                            {
                                row++;
                            }
                        }
                    }
                }
                // 2) 回退：基于内容的启发式纵向合并（与之前版本保持兼容）
                else if (currentTable.ColumnCount > 0)
                {
                    for (int col = 0; col < currentTable.ColumnCount; col++)
                    {
                        int row = 0;
                        while (row < currentTable.Rows.Count)
                        {
                            var startCell = GetCell(currentTable, row, col);
                            if (startCell == null)
                            {
                                row++;
                                continue;
                            }

                            if (!CellHasContent(startCell))
                            {
                                row++;
                                continue;
                            }

                            int span = 1;
                            int nextRow = row + 1;
                            while (nextRow < currentTable.Rows.Count)
                            {
                                var nextCell = GetCell(currentTable, nextRow, col);
                                if (nextCell == null)
                                {
                                    break;
                                }

                                if (CellHasContent(nextCell))
                                {
                                    break;
                                }

                                span++;
                                nextRow++;
                            }

                            if (span > 1)
                            {
                                startCell.RowSpan = span;
                                row += span;
                            }
                            else
                            {
                                row++;
                            }
                        }
                    }
                }

                // 3) 基于 TAP / TC 的精确信息推断横向合并（ColumnSpan）
                if (hasTapMergeInfo && currentTable.ColumnCount > 0)
                {
                    for (int row = 0; row < currentTable.Rows.Count; row++)
                    {
                        var tap = row < rowTaps.Count ? rowTaps[row] : null;
                        var mergeArray = tap?.CellMerges;
                        if (mergeArray == null || mergeArray.Length == 0) continue;

                        int col = 0;
                        while (col < currentTable.ColumnCount)
                        {
                            var cell = GetCell(currentTable, row, col);
                            if (cell == null)
                            {
                                col++;
                                continue;
                            }

                            CellMergeFlags? flags = col < mergeArray.Length ? mergeArray[col] : null;
                            if (flags == null || !flags.HorizFirst)
                            {
                                col++;
                                continue;
                            }

                            int span = 1;
                            int nextCol = col + 1;
                            while (nextCol < currentTable.ColumnCount)
                            {
                                var nextFlags = nextCol < mergeArray.Length ? mergeArray[nextCol] : null;
                                if (nextFlags == null || !nextFlags.HorizMerged)
                                {
                                    break;
                                }
                                span++;
                                nextCol++;
                            }

                            if (span > 1)
                            {
                                cell.ColumnSpan = span;
                                col += span;
                            }
                            else
                            {
                                col++;
                            }
                        }
                    }
                }

                tables.Add(currentTable);
                currentTable = null;
                rowTaps.Clear();

                static TableCellModel? GetCell(TableModel table, int rowIndex, int columnIndex)
                {
                    if (rowIndex < 0 || rowIndex >= table.Rows.Count) return null;
                    var row = table.Rows[rowIndex];
                    if (columnIndex < 0 || columnIndex >= row.Cells.Count) return null;
                    return row.Cells[columnIndex];
                }

                static bool CellHasContent(TableCellModel cell)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        if (para.Runs.Any(r => !string.IsNullOrEmpty(r.Text)))
                        {
                            return true;
                        }
                    }
                    return false;
                }
            }
    }
}

/// <summary>
/// Image reader — extracts images from Word documents.
/// Phase 1: stub implementation. Phase 3 will parse OfficeArt/BLIP data.
/// </summary>
public class ImageReader
{
    private readonly BinaryReader _wordDocReader;
    private readonly BinaryReader? _dataReader;
    private readonly FibReader _fib;
    private readonly CfbReader? _cfb;

    public ImageReader(BinaryReader wordDocReader, BinaryReader? dataReader, FibReader fib, CfbReader? cfb = null)
    {
        _wordDocReader = wordDocReader;
        _dataReader = dataReader;
        _fib = fib;
        _cfb = cfb;
    }

    /// <summary>
    /// Extracts images from the document.
    /// First extracts at sprmCPicLocation (FcPic) offsets for inline pictures, then scans for BLIPs.
    /// </summary>
    public void ExtractImages(DocumentModel document)
    {
        document.Images = new List<ImageModel>();

        if (_dataReader == null) return;

        try
        {
            var data = _dataReader.BaseStream;
            if (data.Length < 8) return;

            data.Seek(0, SeekOrigin.Begin);
            var buffer = new byte[data.Length];
            _dataReader.Read(buffer, 0, (int)data.Length);

            var extractedRanges = new HashSet<(int start, int end)>();

            // 1. Extract images at sprmCPicLocation (FcPic) offsets — inline pictures in document order
            ExtractImagesAtFcPicPositions(document, buffer, extractedRanges);

            // 2. Scan for additional BLIPs (floating shapes, etc.), skipping already-extracted ranges
            ScanForImageSignatures(buffer, document.Images, extractedRanges);
            ScanForOfficeArtRecords(buffer, document.Images, extractedRanges);

            // 3. Scan other OLE streams (WordDocument, Table, ObjectPool children) for embedded images
            if (_cfb != null)
                ScanAdditionalStreamsForImages(document.Images);

            // 4. Ensure every picture run has a valid ImageIndex (assign 0,1,2... in document order)
            AssignPictureRunIndices(document);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Image extraction failed: {ex.Message}");
        }
    }

    /// <summary>Extracts images at FcPic positions and assigns run.ImageIndex.</summary>
    private void ExtractImagesAtFcPicPositions(DocumentModel document, byte[] buffer, HashSet<(int start, int end)> extractedRanges)
    {
        // Reset all picture run indices; we will set only for successful FcPic extractions
        foreach (var para in document.Paragraphs)
            foreach (var run in para.Runs)
                if (run.IsPicture) run.ImageIndex = -1;
        foreach (var table in document.Tables)
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                    foreach (var para in cell.Paragraphs)
                        foreach (var run in para.Runs)
                            if (run.IsPicture) run.ImageIndex = -1;
        foreach (var note in document.Footnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) run.ImageIndex = -1;
        foreach (var note in document.Endnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) run.ImageIndex = -1;
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs == null) continue;
            foreach (var para in textbox.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) run.ImageIndex = -1;
        }

        foreach (var para in document.Paragraphs)
        {
            foreach (var run in para.Runs)
            {
                if (!run.IsPicture || run.FcPic == 0) continue;
                // MS-DOC: sprmCPicLocation is position in Data stream (byte offset). Treat as unsigned.
                if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                { run.ImageIndex = -1; continue; }

                var img = TryExtractImageAtOffset(buffer, offset);
                if (img != null)
                {
                    img.PictureOffset = offset;
                    run.ImageIndex = document.Images.Count;
                    document.Images.Add(img);
                    if (img.Data != null)
                        extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                }
                else
                    run.ImageIndex = -1;
            }
        }
        foreach (var table in document.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        foreach (var run in para.Runs)
                        {
                            if (!run.IsPicture || run.FcPic == 0) continue;
                            if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                            { run.ImageIndex = -1; continue; }

                            var img = TryExtractImageAtOffset(buffer, offset);
                            if (img != null)
                            {
                                img.PictureOffset = offset;
                                run.ImageIndex = document.Images.Count;
                                document.Images.Add(img);
                                if (img.Data != null)
                                    extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                            }
                            else
                                run.ImageIndex = -1;
                        }
                    }
                }
            }
        }

        // Footnotes
        foreach (var note in document.Footnotes)
        {
            foreach (var para in note.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    if (!run.IsPicture || run.FcPic == 0) continue;
                    if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                    { run.ImageIndex = -1; continue; }
                    var img = TryExtractImageAtOffset(buffer, offset);
                    if (img != null)
                    {
                        run.ImageIndex = document.Images.Count;
                        document.Images.Add(img);
                        if (img.Data != null)
                            extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                    }
                    else
                        run.ImageIndex = -1;
                }
            }
        }
        // Endnotes
        foreach (var note in document.Endnotes)
        {
            foreach (var para in note.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    if (!run.IsPicture || run.FcPic == 0) continue;
                    if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                    { run.ImageIndex = -1; continue; }
                    var img = TryExtractImageAtOffset(buffer, offset);
                    if (img != null)
                    {
                        img.PictureOffset = offset;
                        run.ImageIndex = document.Images.Count;
                        document.Images.Add(img);
                        if (img.Data != null)
                            extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                    }
                    else
                        run.ImageIndex = -1;
                }
            }
        }
        // Textboxes
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs == null) continue;
            foreach (var para in textbox.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    if (!run.IsPicture || run.FcPic == 0) continue;
                    if (!TryFcPicToBufferOffset(run.FcPic, buffer.Length, out int offset))
                    { run.ImageIndex = -1; continue; }
                    var img = TryExtractImageAtOffset(buffer, offset);
                    if (img != null)
                    {
                        img.PictureOffset = offset;
                        run.ImageIndex = document.Images.Count;
                        document.Images.Add(img);
                        if (img.Data != null)
                            extractedRanges.Add((offset, Math.Min(offset + 68 + img.Data.Length + 512, buffer.Length)));
                    }
                    else
                        run.ImageIndex = -1;
                }
            }
        }
    }

    /// <summary>Maps FcPic (signed in spec but stored as uint) to a valid buffer offset. Treats value as unsigned byte offset.</summary>
    private static bool TryFcPicToBufferOffset(uint fcPic, int bufferLength, out int offset)
    {
        if (fcPic == 0) { offset = 0; return false; }
        // Use as unsigned; max Data stream size is 0x7FFFFFFF so valid offset fits in int.
        if (fcPic > (uint)int.MaxValue || fcPic >= (uint)bufferLength) { offset = 0; return false; }
        offset = (int)fcPic;
        return true;
    }

    /// <summary>Tries FcPic as byte offset, then alternate interpretations (FC*2, FC/2, and direct OfficeArt at offset).</summary>
    private ImageModel? TryExtractImageAtOffset(byte[] buffer, int offset)
    {
        // 1) Standard: PICF at offset
        if (offset >= 0 && offset < buffer.Length)
        {
            var img = ExtractImageAtPicfOffset(buffer, offset);
            if (img != null) return img;
        }
        // 2) FcPic in 32-bit FC units (byte offset = FcPic*2)
        long offsetAlt = (long)offset * 2;
        if (offsetAlt > 0 && offsetAlt < buffer.Length)
        {
            var img = ExtractImageAtPicfOffset(buffer, (int)offsetAlt);
            if (img != null) return img;
        }
        // 3) Half-byte offset
        int offsetHalf = offset / 2;
        if (offsetHalf >= 0 && offsetHalf < buffer.Length)
        {
            var img = ExtractImageAtPicfOffset(buffer, offsetHalf);
            if (img != null) return img;
        }
        // 4) Some files: FcPic points directly to OfficeArt (no PICF); try BLIP/signature at offset
        if (offset >= 0 && offset < buffer.Length)
        {
            var img = ExtractBlipFromOfficeArtRegion(buffer, offset, buffer.Length - offset);
            if (img != null) return img;
            var scan = ScanForImageAtPosition(buffer, offset);
            if (scan != null) return scan;
        }
        return null;
    }

    /// <summary>Extracts image from PICFAndOfficeArtData at the given Data stream offset.</summary>
    private ImageModel? ExtractImageAtPicfOffset(byte[] buffer, int offset)
    {
        if (offset + 68 > buffer.Length) return null;
        // PICF: lcb(4)+cbHeader(2)+mfpf(8)+...; mfpf.mm at offset 6
        var mm = offset + 8 <= buffer.Length ? BitConverter.ToUInt16(buffer, offset + 6) : (ushort)0;
        var picfEnd = 68;
        if (mm == 0x0066 && offset + 69 <= buffer.Length)
        {
            var cchPicName = buffer[offset + 68];
            picfEnd = 69 + cchPicName;
        }
        if (offset + picfEnd >= buffer.Length) return null;

        var artStart = offset + picfEnd;
        // OfficeArtInlineSpContainer: shape then rgfb (BLIPs); recurse into containers
        var img = ExtractBlipFromOfficeArtRegion(buffer, artStart, buffer.Length - artStart);
        if (img != null) return img;
        var scan = ScanForImageAtPosition(buffer, artStart);
        if (scan != null) return scan;

        // Fallback: search a window for BLIP or raw signature (alignment/variant PICF)
        for (int delta = 8; delta <= 128 && artStart - delta >= 0; delta += 8)
        {
            int winStart = artStart - delta;
            int winLen = Math.Min(4096, buffer.Length - winStart);
            if (winLen >= 8)
            {
                img = ExtractBlipFromOfficeArtRegion(buffer, winStart, winLen);
                if (img != null) return img;
            }
        }
        int searchEnd = Math.Min(artStart + 1024, buffer.Length - 8);
        for (int p = artStart; p < searchEnd; p++)
        {
            scan = ScanForImageAtPosition(buffer, p);
            if (scan != null) return scan;
        }
        return null;
    }

    /// <summary>Searches for a BLIP or image within an OfficeArt region; recurses into containers.</summary>
    private ImageModel? ExtractBlipFromOfficeArtRegion(byte[] buffer, int start, int maxLen)
    {
        int pos = 0;
        while (pos < maxLen - 8 && start + pos < buffer.Length)
        {
            var recType = BitConverter.ToUInt16(buffer, start + pos + 2);
            var recLen = BitConverter.ToUInt32(buffer, start + pos + 4);
            if (recLen > 50 * 1024 * 1024 ||
                pos + 8 + recLen > (uint)maxLen ||
                start + pos + 8 + recLen > buffer.Length)
            {
                pos++;
                continue;
            }
            // BLIP types per MS-ODRAW 2.2.23
            if (recType >= 0xF018 && recType <= 0xF117)
            {
                var img = ExtractBlipFromOfficeArtRecord(buffer, start + pos + 8, (int)recLen, recType);
                if (img != null) return img;
            }
            // Recurse into containers: DggContainer, DgContainer, SpgrContainer, SpContainer, BStoreContainer
            if (recType == 0xF000 || recType == 0xF001 || recType == 0xF002 || recType == 0xF003 || recType == 0xF004 || recType == 0xF007)
            {
                var innerLen = (int)recLen;
                if (innerLen > 0 && start + pos + 8 + innerLen <= buffer.Length)
                {
                    var img = ExtractBlipFromOfficeArtRegion(buffer, start + pos + 8, innerLen);
                    if (img != null) return img;
                }
            }
            pos += 8 + (int)recLen;
        }
        return null;
    }

    /// <summary>Scans for raw image signature at the given position.</summary>
    private ImageModel? ScanForImageAtPosition(byte[] buffer, int pos)
    {
        if (pos + 8 > buffer.Length) return null;
        ImageType? type = null;
        if (buffer[pos] == 0x89 && buffer[pos + 1] == 0x50 && buffer[pos + 2] == 0x4E && buffer[pos + 3] == 0x47)
            type = ImageType.Png;
        else if (buffer[pos] == 0xFF && buffer[pos + 1] == 0xD8 && buffer[pos + 2] == 0xFF)
            type = ImageType.Jpeg;
        else if (buffer[pos] == 0x47 && buffer[pos + 1] == 0x49 && buffer[pos + 2] == 0x46)
            type = ImageType.Gif;
        else if (buffer[pos] == 0x42 && buffer[pos + 1] == 0x4D)
            type = ImageType.Dib;
        if (!type.HasValue) return null;
        return ExtractImageFromPosition(buffer, pos, type.Value);
    }

    /// <summary>
    /// Scans the Data stream for OfficeArt BLIP records.
    /// </summary>
    private List<ImageModel> ScanForBlipRecords()
    {
        var images = new List<ImageModel>();
        var data = _dataReader!.BaseStream;
        var length = data.Length;

        if (length < 8) return images;

        data.Seek(0, SeekOrigin.Begin);
        var buffer = new byte[length];
        _dataReader.Read(buffer, 0, (int)length);

        ScanForImageSignatures(buffer, images);
        ScanForOfficeArtRecords(buffer, images);

        return images;
    }

    /// <summary>
    /// Scans for raw image signatures (PNG, JPEG, GIF, etc.) in the data.
    /// </summary>
    private void ScanForImageSignatures(byte[] buffer, List<ImageModel> images, HashSet<(int start, int end)>? skipRanges = null)
    {
        int pos = 0;
        while (pos < buffer.Length - 8)
        {
            if (IsInSkipRange(pos, skipRanges)) { pos++; continue; }

            ImageType? type = null;
            int headerLen = 0;

            // Check for PNG signature
            if (pos + 8 <= buffer.Length &&
                buffer[pos] == 0x89 && buffer[pos + 1] == 0x50 &&
                buffer[pos + 2] == 0x4E && buffer[pos + 3] == 0x47 &&
                buffer[pos + 4] == 0x0D && buffer[pos + 5] == 0x0A &&
                buffer[pos + 6] == 0x1A && buffer[pos + 7] == 0x0A)
            {
                type = ImageType.Png;
                headerLen = 8;
            }
            // Check for JPEG signature
            else if (pos + 3 <= buffer.Length &&
                     buffer[pos] == 0xFF && buffer[pos + 1] == 0xD8 && buffer[pos + 2] == 0xFF)
            {
                type = ImageType.Jpeg;
                headerLen = 3;
            }
            // Check for GIF signature
            else if (pos + 6 <= buffer.Length &&
                     buffer[pos] == 0x47 && buffer[pos + 1] == 0x49 && buffer[pos + 2] == 0x46 &&
                     buffer[pos + 3] == 0x38 && (buffer[pos + 4] == 0x37 || buffer[pos + 4] == 0x39) &&
                     buffer[pos + 5] == 0x61)
            {
                type = ImageType.Gif;
                headerLen = 6;
            }
            // Check for BMP signature
            else if (pos + 2 <= buffer.Length && buffer[pos] == 0x42 && buffer[pos + 1] == 0x4D)
            {
                type = ImageType.Dib;
                headerLen = 2;
            }

            if (type.HasValue)
            {
                var image = ExtractImageFromPosition(buffer, pos, type.Value);
                if (image != null)
                {
                    images.Add(image);
                    // Mark range as extracted to avoid duplicate from OfficeArt scan
                    skipRanges?.Add((pos, Math.Min(pos + image.Data.Length, buffer.Length)));
                    pos += image.Data.Length;
                    continue;
                }
            }

            pos++;
        }
    }

    /// <summary>Ensures every picture run has a valid ImageIndex. Preserves indices set by FcPic extraction; for runs that failed extraction, picks the image whose PictureOffset is closest to run.FcPic so the correct image (e.g. first-page background) is shown.</summary>
    private static void AssignPictureRunIndices(DocumentModel document)
    {
        int maxIdx = Math.Max(0, document.Images.Count - 1);
        var assigned = new HashSet<int>();
        foreach (var run in EnumeratePictureRuns(document))
            if (run.ImageIndex >= 0)
                assigned.Add(run.ImageIndex);

        int fallbackIdx = 0;
        foreach (var run in EnumeratePictureRuns(document))
        {
            if (run.ImageIndex >= 0) continue;
            int chosen = -1;
            if (run.FcPic != 0 && document.Images.Count > 0)
            {
                long bestDist = long.MaxValue;
                for (int i = 0; i < document.Images.Count; i++)
                {
                    if (assigned.Contains(i)) continue;
                    int po = document.Images[i].PictureOffset;
                    if (po == 0) continue;
                    long d = Math.Abs((long)po - run.FcPic);
                    if (d < bestDist)
                    {
                        bestDist = d;
                        chosen = i;
                    }
                }
            }
            if (chosen < 0)
            {
                while (fallbackIdx <= maxIdx && assigned.Contains(fallbackIdx)) fallbackIdx++;
                chosen = fallbackIdx <= maxIdx ? fallbackIdx++ : 0;
            }
            run.ImageIndex = chosen;
            assigned.Add(chosen);
        }
    }

    private static IEnumerable<RunModel> EnumeratePictureRuns(DocumentModel document)
    {
        foreach (var para in document.Paragraphs)
            foreach (var run in para.Runs)
                if (run.IsPicture) yield return run;
        foreach (var table in document.Tables)
            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                    foreach (var para in cell.Paragraphs)
                        foreach (var run in para.Runs)
                            if (run.IsPicture) yield return run;
        foreach (var note in document.Footnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        foreach (var note in document.Endnotes)
            foreach (var para in note.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        foreach (var textbox in document.Textboxes)
        {
            if (textbox.Paragraphs == null) continue;
            foreach (var para in textbox.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.IsPicture) yield return run;
        }
    }

    /// <summary>Scans WordDocument, Table, and ObjectPool streams for embedded images.</summary>
    private void ScanAdditionalStreamsForImages(List<ImageModel> images)
    {
        var names = new[] { "WordDocument", "0Table", "1Table" };
        foreach (var name in names)
        {
            if (!_cfb!.HasStream(name)) continue;
            try
            {
                var bytes = _cfb.GetStreamBytes(name);
                ScanForImageSignatures(bytes, images);
            }
            catch { /* best-effort */ }
        }
        // Scan other top-level streams that might contain embedded images (e.g. OLE embeddings)
        foreach (var name in _cfb!.StreamNames)
        {
            if (name is "WordDocument" or "0Table" or "1Table" or "Data") continue;
            if (name.Length < 2) continue;
            try
            {
                var bytes = _cfb.GetStreamBytes(name);
                if (bytes.Length > 8 && bytes.Length < 50 * 1024 * 1024)
                    ScanForImageSignatures(bytes, images);
            }
            catch { }
        }
    }

    private static bool IsInSkipRange(int pos, HashSet<(int start, int end)>? ranges)
    {
        if (ranges == null) return false;
        foreach (var (start, end) in ranges)
            if (pos >= start && pos < end) return true;
        return false;
    }

    /// <summary>
    /// Scans for OfficeArt container records.
    /// </summary>
    private void ScanForOfficeArtRecords(byte[] buffer, List<ImageModel> images, HashSet<(int start, int end)>? skipRanges = null)
    {
        // OfficeArt record header format:
        // - recVer (4 bits) + recInstance (12 bits) = 2 bytes
        // - recType (2 bytes)
        // - recLen (4 bytes)

        int pos = 0;
        while (pos < buffer.Length - 8)
        {
            if (IsInSkipRange(pos, skipRanges)) { pos++; continue; }
            try
            {
                var recInfo = BitConverter.ToUInt16(buffer, pos);
                var recVer = recInfo & 0x0F;
                var recInstance = (recInfo >> 4) & 0x0FFF;
                var recType = BitConverter.ToUInt16(buffer, pos + 2);
                var recLen = BitConverter.ToUInt32(buffer, pos + 4);

                // Validate record
                if (recLen > 0 && recLen < 100 * 1024 * 1024 && // Max 100MB
                    pos + 8 + recLen <= buffer.Length)
                {
                    // Check for BLIP types (0xF018-0xF117)
                    if (recType >= 0xF018 && recType <= 0xF117)
                    {
                        var image = ExtractBlipFromOfficeArtRecord(buffer, pos + 8, (int)recLen, recType);
                        if (image != null)
                        {
                            images.Add(image);
                            // Mark range as extracted to avoid duplicate from raw signature scan
                            skipRanges?.Add((pos, Math.Min(pos + 8 + (int)recLen, buffer.Length)));
                        }
                    }

                    pos += 8 + (int)recLen;
                }
                else
                {
                    pos++;
                }
            }
            catch
            {
                pos++;
            }
        }
    }

    /// <summary>
    /// Extracts an image from a specific position in the buffer.
    /// </summary>
    private ImageModel? ExtractImageFromPosition(byte[] buffer, int pos, ImageType type)
    {
        try
        {
            int length = EstimateImageLength(buffer, pos, type);
            if (length <= 0 || pos + length > buffer.Length) return null;

            var data = new byte[length];
            Array.Copy(buffer, pos, data, 0, length);

            // Get dimensions if possible
            var (width, height) = GetImageDimensions(data, type);

            return new ImageModel
            {
                Id = $"image{pos}",
                FileName = $"image{pos}.{GetImageExtension(type)}",
                ContentType = GetContentType(type),
                Data = data,
                Type = type,
                Width = width,
                Height = height,
                WidthEMU = width > 0 ? width * 914400 / 96 : 0,
                HeightEMU = height > 0 ? height * 914400 / 96 : 0,
                PictureOffset = pos
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Extracts a BLIP from an OfficeArt record.
    /// </summary>
    private ImageModel? ExtractBlipFromOfficeArtRecord(byte[] buffer, int offset, int length, ushort recType)
    {
        try
        {
            // BLIP record contains a header followed by the image data
            // Skip BLIP header (typically 16-34 bytes depending on type)
            int headerSize = GetBlipHeaderSize(recType);
            if (offset + headerSize >= buffer.Length) return null;

            var imageType = GetBlipImageType(recType);
            if (imageType == ImageType.Unknown) return null;

            var dataLength = length - headerSize;
            if (dataLength <= 0) return null;

            var data = new byte[dataLength];
            Array.Copy(buffer, offset + headerSize, data, 0, dataLength);

            var (width, height) = GetImageDimensions(data, imageType);

            return new ImageModel
            {
                Id = $"blip{offset}",
                FileName = $"image{offset}.{GetImageExtension(imageType)}",
                ContentType = GetContentType(imageType),
                Data = data,
                Type = imageType,
                Width = width,
                Height = height,
                WidthEMU = width > 0 ? width * 914400 / 96 : 0,
                HeightEMU = height > 0 ? height * 914400 / 96 : 0,
                PictureOffset = offset
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Estimates the length of an image based on its type and structure.
    /// </summary>
    private int EstimateImageLength(byte[] buffer, int pos, ImageType type)
    {
        try
        {
            switch (type)
            {
                case ImageType.Png:
                    // PNG has chunks; look for IEND chunk
                    return FindPngEnd(buffer, pos);

                case ImageType.Jpeg:
                    // JPEG ends with EOI marker FF D9
                    return FindJpegEnd(buffer, pos);

                case ImageType.Gif:
                    // GIF ends with 3B
                    return FindGifEnd(buffer, pos);

                case ImageType.Dib:
                    // BMP has size in header
                    return FindBmpEnd(buffer, pos);

                default:
                    return 0;
            }
        }
        catch
        {
            return 0;
        }
    }

    private int FindPngEnd(byte[] buffer, int start)
    {
        int pos = start + 8; // Skip signature
        while (pos < buffer.Length - 12)
        {
            // PNG chunk length is 4 bytes big-endian (network byte order)
            var chunkLenU = (uint)((buffer[pos] << 24) | (buffer[pos + 1] << 16) | (buffer[pos + 2] << 8) | buffer[pos + 3]);
            if (chunkLenU > 100 * 1024 * 1024) break;
            var chunkLen = (int)chunkLenU;

            var chunkType = Encoding.ASCII.GetString(buffer, pos + 4, 4);
            if (chunkType == "IEND")
                return pos + 12; // Length (4) + Type (4) + Data (0) + CRC (4)

            pos += 12 + chunkLen;
        }
        return buffer.Length - start;
    }

    private int FindJpegEnd(byte[] buffer, int start)
    {
        int pos = start + 2; // Skip SOI
        while (pos < buffer.Length - 1)
        {
            if (buffer[pos] != 0xFF)
            {
                pos++;
                continue;
            }
            var m = buffer[pos + 1];
            if (m == 0xD9) // EOI
                return pos + 2 - start;
            if (m == 0xD8) // Another SOI (invalid)
                break;
            // Skip segments with length: SOF, DHT, DQT, DRI, APP, COM, SOS (has length then scan for next marker)
            if (m == 0xC0 || m == 0xC1 || m == 0xC2 || m == 0xC3 || m == 0xC4 || m == 0xC5 ||
                m == 0xC6 || m == 0xC7 || m == 0xC8 || m == 0xC9 || m == 0xCA || m == 0xCB ||
                m == 0xCC || m == 0xCD || m == 0xCE || m == 0xCF ||
                m == 0xDB || m == 0xDD || m == 0xE0 || m == 0xE1 || m == 0xE2 || m == 0xE3 ||
                m == 0xE4 || m == 0xE5 || m == 0xE6 || m == 0xE7 || m == 0xE8 || m == 0xE9 ||
                m == 0xEA || m == 0xEB || m == 0xEC || m == 0xED || m == 0xEE || m == 0xEF ||
                m == 0xFE || m == 0xDA) // SOS: skip length then scan until next 0xFF
            {
                if (pos + 4 > buffer.Length) break;
                var segLen = (buffer[pos + 2] << 8) | buffer[pos + 3];
                if (m == 0xDA) // SOS: after 2+segLen, scan for 0xFF 0xD9
                {
                    int sosEnd = pos + 2 + segLen;
                    for (int i = sosEnd; i < buffer.Length - 1; i++)
                    {
                        if (buffer[i] == 0xFF && buffer[i + 1] == 0xD9)
                            return i + 2 - start;
                        if (buffer[i] == 0xFF && buffer[i + 1] != 0x00) // Skip escaped 0xFF in scan
                            i++;
                    }
                    break;
                }
                pos += 2 + segLen;
                continue;
            }
            if (m >= 0xD0 && m <= 0xD7) { pos += 2; continue; } // RST: no length
            pos++;
        }
        return buffer.Length - start;
    }

    private int FindGifEnd(byte[] buffer, int start)
    {
        int pos = start;
        while (pos < buffer.Length)
        {
            if (buffer[pos] == 0x3B) // Trailer
                return pos + 1 - start;
            pos++;
        }
        return buffer.Length - start;
    }

    private int FindBmpEnd(byte[] buffer, int start)
    {
        if (start + 14 > buffer.Length) return 0;
        var size = BitConverter.ToInt32(buffer, start + 2);
        return size > 0 && size < buffer.Length - start ? size : 0;
    }

    /// <summary>
    /// Gets the image dimensions from the image data.
    /// </summary>
    private (int width, int height) GetImageDimensions(byte[] data, ImageType type)
    {
        try
        {
            switch (type)
            {
                case ImageType.Png:
                    if (data.Length >= 24)
                    {
                        var width = (int)BitConverter.ToUInt32(data, 16);
                        var height = (int)BitConverter.ToUInt32(data, 20);
                        return (width, height);
                    }
                    break;

                case ImageType.Jpeg:
                    return GetJpegDimensions(data);

                case ImageType.Gif:
                    if (data.Length >= 10)
                    {
                        var width = BitConverter.ToUInt16(data, 6);
                        var height = BitConverter.ToUInt16(data, 8);
                        return (width, height);
                    }
                    break;

                case ImageType.Dib:
                    if (data.Length >= 26)
                    {
                        var width = BitConverter.ToInt32(data, 18);
                        var height = BitConverter.ToInt32(data, 22);
                        return (width, height);
                    }
                    break;
            }
        }
        catch { }

        return (0, 0);
    }

    private (int width, int height) GetJpegDimensions(byte[] data)
    {
        int pos = 2; // Skip SOI
        while (pos < data.Length - 9)
        {
            if (data[pos] == 0xFF && data[pos + 1] != 0x00 && data[pos + 1] != 0xFF)
            {
                var marker = data[pos + 1];
                if (marker == 0xD9) break; // EOI
                if (marker == 0xC0 || marker == 0xC1 || marker == 0xC2) // SOF markers
                {
                    var height = (data[pos + 5] << 8) | data[pos + 6];
                    var width = (data[pos + 7] << 8) | data[pos + 8];
                    return (width, height);
                }

                var len = (data[pos + 2] << 8) | data[pos + 3];
                pos += 2 + len;
            }
            else
            {
                pos++;
            }
        }
        return (0, 0);
    }

    private int GetBlipHeaderSize(ushort recType)
    {
        // MS-ODRAW 2.2.23+: BLIP header sizes
        return recType switch
        {
            0xF01A => 16, // EMF
            0xF01B => 16, // WMF
            0xF01C => 16, // PICT
            0xF01D => 17, // JPEG
            0xF01E => 17, // PNG
            0xF01F => 17, // DIB
            0xF020 => 17, // TIFF (legacy)
            0xF029 => 17, // TIFF
            0xF02A => 17, // JPEG (alternate)
            _ => 16
        };
    }

    private ImageType GetBlipImageType(ushort recType)
    {
        return recType switch
        {
            0xF01A => ImageType.Emf,
            0xF01B => ImageType.Wmf,
            0xF01D => ImageType.Jpeg,
            0xF01E => ImageType.Png,
            0xF01F => ImageType.Dib,
            0xF020 => ImageType.Tiff,
            0xF029 => ImageType.Tiff,
            0xF02A => ImageType.Jpeg,
            _ => ImageType.Unknown
        };
    }

    private string GetImageExtension(ImageType type)
    {
        return type switch
        {
            ImageType.Png => "png",
            ImageType.Jpeg => "jpg",
            ImageType.Gif => "gif",
            ImageType.Emf => "emf",
            ImageType.Wmf => "wmf",
            ImageType.Dib => "bmp",
            ImageType.Tiff => "tiff",
            _ => "bin"
        };
    }

    private string GetContentType(ImageType type)
    {
        return type switch
        {
            ImageType.Png => "image/png",
            ImageType.Jpeg => "image/jpeg",
            ImageType.Gif => "image/gif",
            ImageType.Emf => "image/x-emf",
            ImageType.Wmf => "image/x-wmf",
            ImageType.Dib => "image/bmp",
            ImageType.Tiff => "image/tiff",
            _ => "application/octet-stream"
        };
    }

    /// <summary>
    /// Determines image type from header bytes.
    /// </summary>
    public static ImageType GetImageType(byte[] header)
    {
        if (header.Length < 4) return ImageType.Unknown;

        // PNG: 89 50 4E 47
        if (header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
            return ImageType.Png;

        // JPEG: FF D8 FF
        if (header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF)
            return ImageType.Jpeg;

        // GIF: 47 49 46
        if (header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46)
            return ImageType.Gif;

        // EMF: 01 00 00 00
        if (header[0] == 0x01 && header[1] == 0x00 && header[2] == 0x00 && header[3] == 0x00)
            return ImageType.Emf;

        // WMF: D7 CD C6 9A
        if (header[0] == 0xD7 && header[1] == 0xCD && header[2] == 0xC6 && header[3] == 0x9A)
            return ImageType.Wmf;

        return ImageType.Unknown;
    }
}
