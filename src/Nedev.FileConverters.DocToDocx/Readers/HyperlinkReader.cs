using System.Text;
using System.Text.RegularExpressions;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Hyperlink reader - extracts hyperlinks from field codes.
///
/// In Word 97-2003, hyperlinks are stored as field codes:
///   HYPERLINK "url" [switches]
///
/// The field structure is:
///   - Field start (19)
///   - Field code (HYPERLINK ...)
///   - Field separator (20)
///   - Display text
///   - Field end (21)
/// </summary>
public class HyperlinkReader
{
    // Regex to match HYPERLINK field codes
    private static readonly Regex HyperlinkRegex = new(
        @"HYPERLINK\s+""([^""]+)""(?:\s+\\l\s+""([^""]+)"")?(?:\s+\\m\s+""([^""]+)"")?",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HyperlinkBookmarkOnlyRegex = new(
        @"HYPERLINK\s+\\l\s+""([^""]+)""(?:\s+\\m\s+""([^""]+)"")?",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HyperlinkSimpleRegex = new(
        @"HYPERLINK\s+(?:""([^""]+)""|'([^']+)'|(\S+))",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public List<HyperlinkModel> Hyperlinks { get; private set; } = new();

    /// <summary>
    /// Parses hyperlinks from a field code string.
    /// </summary>
    public HyperlinkModel? ParseHyperlink(string fieldCode)
    {
        if (string.IsNullOrWhiteSpace(fieldCode))
            return null;

        // Try regex match
        var match = HyperlinkRegex.Match(fieldCode);
        bool bookmarkOnly = false;
        bool simpleMatch = false;
        if (!match.Success)
        {
            match = HyperlinkBookmarkOnlyRegex.Match(fieldCode);
            bookmarkOnly = match.Success;
        }

        if (!match.Success)
        {
            match = HyperlinkSimpleRegex.Match(fieldCode);
            simpleMatch = match.Success;
        }

        if (!match.Success)
            return null;

        string url = string.Empty;
        string? bookmark = null;

        if (bookmarkOnly)
        {
            bookmark = match.Groups[1].Value;
        }
        else
        {
            url = match.Groups[1].Value;
            if (string.IsNullOrEmpty(url) && match.Groups.Count > 2)
            {
                url = match.Groups[2].Value;
            }
            if (string.IsNullOrEmpty(url) && match.Groups.Count > 3)
            {
                url = match.Groups[3].Value;
            }

            if (!simpleMatch)
            {
                bookmark = match.Groups.Count > 2 && match.Groups[2].Success ? match.Groups[2].Value : null;
            }
        }

        if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(bookmark))
            return null;

        NormalizeTarget(ref url, ref bookmark);

        return new HyperlinkModel
        {
            Url = url,
            Bookmark = bookmark,
            IsExternal = !string.IsNullOrEmpty(url)
        };
    }

    /// <summary>
    /// Checks if a field code represents a hyperlink.
    /// </summary>
    public bool IsHyperlinkField(string fieldCode)
    {
        return !string.IsNullOrWhiteSpace(fieldCode) &&
               fieldCode.TrimStart().StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Extracts the display text from hyperlink field code if present.
    /// </summary>
    public string GetDisplayText(string fieldCode, string defaultText)
    {
        // If field code has \o switch, it specifies display text
        var match = Regex.Match(fieldCode, @"\\o\s+""([^""]+)""", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            return match.Groups[1].Value;
        }

        return defaultText;
    }

    /// <summary>
    /// Creates a hyperlink model from a URL string.
    /// </summary>
    public HyperlinkModel CreateHyperlink(string url, string? displayText = null)
    {
        string normalizedUrl = url;
        string? bookmark = null;
        NormalizeTarget(ref normalizedUrl, ref bookmark);

        return new HyperlinkModel
        {
            Url = normalizedUrl,
            Bookmark = bookmark,
            DisplayText = displayText,
            IsExternal = normalizedUrl.StartsWith("http://") ||
                        normalizedUrl.StartsWith("https://") ||
                        normalizedUrl.StartsWith("ftp://") ||
                        normalizedUrl.StartsWith("mailto:") ||
                        normalizedUrl.StartsWith("file://")
        };
    }

    private static void NormalizeTarget(ref string url, ref string? bookmark)
    {
        if (!string.IsNullOrEmpty(url) && url.StartsWith("#", StringComparison.Ordinal))
        {
            bookmark ??= url.Substring(1);
            url = string.Empty;
            return;
        }

        if (string.IsNullOrEmpty(url))
            return;

        var hashIndex = url.IndexOf('#');
        if (hashIndex < 0)
            return;

        if (hashIndex + 1 < url.Length)
            bookmark ??= url.Substring(hashIndex + 1);

        url = url.Substring(0, hashIndex);
    }

    /// <summary>
    /// Detects URLs in plain text and converts them to hyperlinks.
    /// </summary>
    public List<HyperlinkModel> DetectUrls(string text)
    {
        var links = new List<HyperlinkModel>();

        // Simple URL detection regex
        var urlRegex = new Regex(
            @"(https?://|ftp://|mailto:)[^\s<>""]+",
            RegexOptions.Compiled);

        var matches = urlRegex.Matches(text);
        foreach (Match match in matches)
        {
            links.Add(new HyperlinkModel
            {
                Url = match.Value,
                IsExternal = true
            });
        }

        return links;
    }
}


