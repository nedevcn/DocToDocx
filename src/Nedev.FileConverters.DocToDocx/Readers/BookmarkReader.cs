using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Bookmark reader - parses PlcfBkf (bookmark first) and PlcfBkl (bookmark last) structures
/// from the Table stream.
///
/// Based on MS-DOC specification §2.7.
///
/// Bookmarks in Word are stored as:
///   - PlcfBkf: Array of CPs marking bookmark starts + BFK structures
///   - PlcfBkl: Array of CPs marking bookmark ends
///   - SttbfBkmk: String table with bookmark names
/// </summary>
public class BookmarkReader
{
    private readonly BinaryReader _tableReader;
    private readonly BinaryReader _wordDocReader;
    private readonly FibReader _fib;

    public List<BookmarkModel> Bookmarks { get; private set; } = new();

    public BookmarkReader(
        BinaryReader tableReader,
        BinaryReader wordDocReader,
        FibReader fib)
    {
        _tableReader = tableReader;
        _wordDocReader = wordDocReader;
        _fib = fib;
    }

    /// <summary>
    /// Reads bookmarks from the document.
    /// </summary>
    public void Read()
    {
        if (_fib.FcPlcfBkf == 0 || _fib.LcbPlcfBkf == 0)
        {
            // No bookmarks
            return;
        }

        if (!_tableReader.CanReadRange(_fib.FcPlcfBkf, _fib.LcbPlcfBkf))
        {
            Logger.Warning($"Skipped bookmarks because PlcfBkf range 0x{_fib.FcPlcfBkf:X}/0x{_fib.LcbPlcfBkf:X} exceeds the Table stream.");
            return;
        }

        try
        {
            ReadPlcfBkf();
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to read bookmarks", ex);
        }
    }

    /// <summary>
    /// Reads PlcfBkf structure (bookmark first positions).
    ///
    /// Structure:
    ///   - Array of CPs (n+1 entries)
    ///   - Array of BKF structures (n entries)
    ///
    /// BKF structure:
    ///   - ibkl (2 bytes) - index into PlcfBkl
    ///   - bkf_flags (2 bytes) - flags
    ///   - lTag (4 bytes) - bookmark tag (if complex)
    ///   - additional data for complex bookmarks
    /// </summary>
    private void ReadPlcfBkf()
    {
        if (_fib.FcPlcfBkf == 0 || _fib.LcbPlcfBkf == 0) return;

        if (_fib.LcbPlcfBkf < 12)
        {
            Logger.Warning($"Skipped bookmarks because PlcfBkf length 0x{_fib.LcbPlcfBkf:X} is too small to contain bookmark data.");
            return;
        }

        _tableReader.BaseStream.Seek(_fib.FcPlcfBkf, SeekOrigin.Begin);
        var fcEndBkf = _fib.FcPlcfBkf + _fib.LcbPlcfBkf;

        // PlcfBkf: (n+1) CPs followed by n BKF structures (each 4 bytes minimum)
        // BKF size is 4 for Word 97, but can be larger in later versions. 
        // We'll assume 4 bytes (ibkl: 2, bkf: 2)
        int n = (int)((_fib.LcbPlcfBkf - 4) / 8); 
        if (n <= 0) return;

        int expectedBkfBytes = (n + 1) * sizeof(int) + n * sizeof(uint);
        if (expectedBkfBytes > _fib.LcbPlcfBkf)
        {
            Logger.Warning($"Skipped bookmarks because PlcfBkf length 0x{_fib.LcbPlcfBkf:X} is inconsistent with {n} bookmark entries.");
            return;
        }

        var startCps = new int[n];
        for (int i = 0; i < n; i++) startCps[i] = _tableReader.ReadInt32();
        _tableReader.ReadInt32(); // Boundary CP

        var ibkls = new ushort[n];
        for (int i = 0; i < n; i++)
        {
            ibkls[i] = _tableReader.ReadUInt16();
            _tableReader.ReadUInt16(); // bkf flags (unused for now)
        }

        // Read End CPs from PlcfBkl
        var endCps = ReadPlcfBkl(n);
        
        // Read Names from SttbfBkmk
        var names = ReadSttbfBkmk(n);

        for (int i = 0; i < n; i++)
        {
            var bookmark = new BookmarkModel
            {
                Index = i,
                StartCp = startCps[i],
                EndCp = (ibkls[i] < endCps.Count) ? endCps[ibkls[i]] : startCps[i],
                Name = (i < names.Count) ? names[i] : $"Bookmark_{i}"
            };
            Bookmarks.Add(bookmark);
        }
    }

    private List<int> ReadPlcfBkl(int n)
    {
        var endCps = new List<int>();
        if (_fib.FcPlcfBkl == 0 || _fib.LcbPlcfBkl == 0) return endCps;

        long expectedBytes = (n + 1L) * sizeof(int);
        if (!_tableReader.CanReadRange(_fib.FcPlcfBkl, _fib.LcbPlcfBkl) || _fib.LcbPlcfBkl < expectedBytes)
        {
            Logger.Warning($"Skipped bookmark end positions because PlcfBkl range 0x{_fib.FcPlcfBkl:X}/0x{_fib.LcbPlcfBkl:X} is inconsistent with {n} bookmark entries.");
            return endCps;
        }

        _tableReader.BaseStream.Seek(_fib.FcPlcfBkl, SeekOrigin.Begin);
        for (int i = 0; i <= n; i++)
        {
            endCps.Add(_tableReader.ReadInt32());
        }
        return endCps;
    }

    private List<string> ReadSttbfBkmk(int n)
    {
        var names = SttbfHelper.ReadSttbf(_tableReader, _fib.FcSttbfBkmk, _fib.LcbSttbfBkmk);
        if (names.Count > n)
            return names.Take(n).ToList();

        return names;
    }
    
    /// <summary>
    /// Gets bookmark at a specific character position.
    /// </summary>
    public BookmarkModel? GetBookmarkAtCp(int cp)
    {
        return Bookmarks.FirstOrDefault(b => b.StartCp <= cp && b.EndCp > cp);
    }

    /// <summary>
    /// Checks if a run is part of a bookmark.
    /// </summary>
    public bool IsInBookmark(RunModel run, out BookmarkModel? bookmark)
    {
        bookmark = GetBookmarkAtCp(run.CharacterPosition);
        return bookmark != null;
    }
}


