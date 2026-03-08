using System.IO;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Reads the File Information Block (FIB) from a Word 97-2003 binary document.
/// The FIB contains the main document metadata and pointers (offsets and lengths)
/// to all other structures in the document.
/// </summary>
public class FibReader
{
    private const int IndexStshf = 1;
    private const int IndexPlcffndRef = 2;
    private const int IndexPlcffndTxt = 3;
    private const int IndexPlcfandRef = 4;
    private const int IndexPlcfandTxt = 5;
    private const int IndexPlcfSed = 6;
    private const int IndexPlcfHdd = 11;
    private const int IndexPlcfBteChpx = 12;
    private const int IndexPlcfBtePapx = 13;
    private const int IndexSttbfFfn = 15;
    private const int IndexPlcfFldMom = 16;
    private const int IndexPlcfFldHdr = 17;
    private const int IndexPlcfFldFtn = 18;
    private const int IndexPlcfFldAtn = 19;
    private const int IndexSttbfBkmk = 21;
    private const int IndexPlcfBkf = 22;
    private const int IndexPlcfBkl = 23;
    private const int IndexDop = 31;
    private const int IndexClx = 33;
    private const int IndexPlcSpaMom = 40;
    private const int IndexPlcfAtnbkf = 42;
    private const int IndexPlcfAtnbkl = 43;
    private const int IndexPlcfendRef = 46;
    private const int IndexPlcfendTxt = 47;
    private const int IndexPlcfFldEdn = 48;
    private const int IndexSttbfRgtlv = 51;
    private const int IndexTxbxText = 56;
    private const int IndexPlcfFldTxbx = 57;
    private const int IndexPlcfLst = 73;
    private const int IndexPlfLfo = 74;

    private readonly BinaryReader _reader;

    // ── FibBase fields ──────────────────────────────────────────────
    public ushort WIdent { get; private set; }
    public ushort NFib { get; private set; }
    public ushort Unused { get; private set; }
    public ushort Lid { get; private set; }
    public ushort PnNext { get; private set; }
    public bool FDot { get; private set; }
    public bool FGlsy { get; private set; }
    public bool FComplex { get; private set; }
    public bool FHasPic { get; private set; }
    public ushort CQuickSaves { get; private set; }
    public bool FEncrypted { get; private set; }
    public bool FWhichTblStm { get; private set; }
    public bool FExtChar { get; private set; }
    public bool FObfuscated { get; private set; }
    public ushort NFibBack { get; private set; }
    public uint LKey { get; private set; }
    public byte Envr { get; private set; }
    public bool FMac { get; private set; }
    public bool FEmptySpecial { get; private set; }
    public bool FLoadOverridePage { get; private set; }
    public uint FcMin { get; private set; }
    public uint FcMac { get; private set; }

    // ── FibRgLw fields (CP counts) ──────────────────────────────────
    public int CcpText { get; private set; }
    public int CcpFtn { get; private set; }
    public int CcpHdd { get; private set; }
    public int CcpAtn { get; private set; }
    public int CcpEdn { get; private set; }
    public int CcpTxbx { get; private set; }
    public int CcpHdrTxbx { get; private set; }

    // ── FibRgFcLcb fields (Offsets and Lengths) ─────────────────────
    public uint FcStshf { get; private set; }
    public uint LcbStshf { get; private set; }
    public uint FcPlcfSed { get; private set; }
    public uint LcbPlcfSed { get; private set; }
    public uint FcPlcfBteChpx { get; private set; }
    public uint LcbPlcfBteChpx { get; private set; }
    public uint FcPlcfBtePapx { get; private set; }
    public uint LcbPlcfBtePapx { get; private set; }
    public uint FcPlcfFldMom { get; private set; }
    public uint LcbPlcfFldMom { get; private set; }
    public uint FcPlcffndRef { get; private set; }
    public uint LcbPlcffndRef { get; private set; }
    public uint FcPlcffndTxt { get; private set; }
    public uint LcbPlcffndTxt { get; private set; }
    public uint FcPlcfandRef { get; private set; }
    public uint LcbPlcfandRef { get; private set; }
    public uint FcPlcfandTxt { get; private set; }
    public uint LcbPlcfandTxt { get; private set; }
    public uint FcPlcfBkf { get; private set; }
    public uint LcbPlcfBkf { get; private set; }
    public uint FcPlcfBkl { get; private set; }
    public uint LcbPlcfBkl { get; private set; }
    public uint FcSttbfAtnMod { get; private set; }
    public uint LcbSttbfAtnMod { get; private set; }
    public uint FcPlcfAtnbkf { get; private set; }
    public uint LcbPlcfAtnbkf { get; private set; }
    public uint FcPlcfAtnbkl { get; private set; }
    public uint LcbPlcfAtnbkl { get; private set; }
    public uint FcPlcfFldAtn { get; private set; }
    public uint LcbPlcfFldAtn { get; private set; }
    public uint FcPlcfFldEdn { get; private set; }
    public uint LcbPlcfFldEdn { get; private set; }
    public uint FcPlcfFldFtn { get; private set; }
    public uint LcbPlcfFldFtn { get; private set; }
    public uint FcPlcfFldHdr { get; private set; }
    public uint LcbPlcfFldHdr { get; private set; }
    public uint FcPlcfFldTxbx { get; private set; }
    public uint LcbPlcfFldTxbx { get; private set; }
    public uint FcSttbfBkmk { get; private set; }
    public uint LcbSttbfBkmk { get; private set; }
    public uint FcPlcfHdd { get; private set; }
    public uint LcbPlcfHdd { get; private set; }
    public uint FcClx { get; private set; }
    public uint LcbClx { get; private set; }
    public uint FcPlcSpaMom { get; private set; }
    public uint LcbPlcSpaMom { get; private set; }
    public uint FcPlcfendRef { get; private set; }
    public uint LcbPlcfendRef { get; private set; }
    public uint FcPlcfendTxt { get; private set; }
    public uint LcbPlcfendTxt { get; private set; }
    public uint FcFtn { get; private set; }
    public uint LcbFtn { get; private set; }
    public uint FcEnd { get; private set; }
    public uint LcbEnd { get; private set; }
    public uint FcAnot { get; private set; }
    public uint LcbAnot { get; private set; }
    public uint FcTxbx { get; private set; }
    public uint LcbTxbx { get; private set; }
    public uint FcGlsy { get; private set; }
    public uint LcbGlsy { get; private set; }
    public uint FcData { get; private set; }
    public uint LcbData { get; private set; }
    public uint FcPlcfLst { get; private set; }
    public uint LcbPlcfLst { get; private set; }
    public uint FcPlfLfo { get; private set; }
    public uint LcbPlfLfo { get; private set; }
    public uint FcSttbfFfn { get; private set; }
    public uint LcbSttbfFfn { get; private set; }
    public uint FcDop { get; private set; }
    public uint LcbDop { get; private set; }
    public uint FcSttbfRgtlv { get; private set; }
    public uint LcbSttbfRgtlv { get; private set; }

    // Legacy Aliases for compatibility with older code
    public uint StshOffset => FcStshf;
    public bool IsComplex => FComplex;
    public uint DopOffset => FcDop;
    public uint PnFbpClx => FcClx;
    public uint TextBaseOffset => 0; // In standard implementations, this is often treated as 0 relative to stream

    /// <summary>
    /// Name of the Table stream to use ("0Table" or "1Table")
    /// </summary>
    public string TableStreamName => FWhichTblStm ? "1Table" : "0Table";

    private readonly List<(uint fc, uint lcb)> _rgFcLcb = new();

    public FibReader(BinaryReader reader)
    {
        _reader = reader;
    }

    public void Read()
    {
        _reader.BaseStream.Seek(0, SeekOrigin.Begin);
        ReadFibBase();
        ReadFibRgW();
        ReadFibRgLw();
        ReadFibRgFcLcb();
    }

    private void ReadFibBase()
    {
        WIdent = _reader.ReadUInt16();
        if (WIdent != WordConsts.FIB_MAGIC_NUMBER && WIdent != WordConsts.FIB_MAGIC_NUMBER_OLD)
            throw new InvalidDataException($"Invalid magic: 0x{WIdent:X4}");

        NFib = _reader.ReadUInt16();
        Unused = _reader.ReadUInt16();
        Lid = _reader.ReadUInt16();
        PnNext = _reader.ReadUInt16();

        var flagsA = _reader.ReadUInt16();
        FDot         = (flagsA & 0x01) != 0;
        FGlsy        = (flagsA & 0x02) != 0;
        FComplex     = (flagsA & 0x04) != 0;
        FHasPic      = (flagsA & 0x08) != 0;
        CQuickSaves  = (ushort)((flagsA >> 4) & 0x0F);
        FEncrypted   = (flagsA & 0x100) != 0;
        FWhichTblStm = (flagsA & 0x200) != 0;
        FExtChar     = (flagsA & 0x1000) != 0;
        FObfuscated  = (flagsA & 0x8000) != 0;

        NFibBack = _reader.ReadUInt16();
        LKey = _reader.ReadUInt32();
        Envr = _reader.ReadByte();

        var flagsB = _reader.ReadByte();
        FMac               = (flagsB & 0x01) != 0;
        FEmptySpecial      = (flagsB & 0x02) != 0;
        FLoadOverridePage  = (flagsB & 0x04) != 0;

        _reader.ReadUInt16(); // reserved
        _reader.ReadUInt16(); // reserved
        FcMin = _reader.ReadUInt32();
        FcMac = _reader.ReadUInt32();
    }

    private void ReadFibRgW()
    {
        var csw = _reader.ReadUInt16();
        for (int i = 0; i < csw; i++) _reader.ReadUInt16();
    }

    private void ReadFibRgLw()
    {
        var cslw = _reader.ReadUInt16();
        var rglw = new int[cslw];
        for (int i = 0; i < cslw; i++) rglw[i] = _reader.ReadInt32();

        if (cslw > 3)  CcpText    = rglw[3];
        if (cslw > 4)  CcpFtn     = rglw[4];
        if (cslw > 5)  CcpHdd     = rglw[5];
        if (cslw > 7)  CcpAtn     = rglw[7];
        if (cslw > 8)  CcpEdn     = rglw[8];
        if (cslw > 9)  CcpTxbx    = rglw[9];
        if (cslw > 10) CcpHdrTxbx = rglw[10];
    }

    private void ReadFibRgFcLcb()
    {
        var cbRgFcLcb = _reader.ReadUInt16();
        _rgFcLcb.Clear();
        for (int i = 0; i < cbRgFcLcb; i++)
        {
            var fc = _reader.ReadUInt32();
            var lcb = _reader.ReadUInt32();
            _rgFcLcb.Add((fc, lcb));
        }

        (FcStshf, LcbStshf)             = GetFcLcb(IndexStshf);
        (FcPlcffndRef, LcbPlcffndRef)   = GetFcLcb(IndexPlcffndRef);
        (FcPlcffndTxt, LcbPlcffndTxt)   = GetFcLcb(IndexPlcffndTxt);
        (FcPlcfandRef, LcbPlcfandRef)   = GetFcLcb(IndexPlcfandRef);
        (FcPlcfandTxt, LcbPlcfandTxt)   = GetFcLcb(IndexPlcfandTxt);
        (FcPlcfSed, LcbPlcfSed)         = GetFcLcb(IndexPlcfSed);
        (FcPlcfHdd, LcbPlcfHdd)         = GetFcLcb(IndexPlcfHdd);
        (FcPlcfBteChpx, LcbPlcfBteChpx) = GetFcLcb(IndexPlcfBteChpx);
        (FcPlcfBtePapx, LcbPlcfBtePapx) = GetFcLcb(IndexPlcfBtePapx);
        (FcSttbfFfn, LcbSttbfFfn)       = GetFcLcb(IndexSttbfFfn);
        (FcPlcfFldMom, LcbPlcfFldMom)   = GetFcLcb(IndexPlcfFldMom);
        (FcPlcfFldHdr, LcbPlcfFldHdr)   = GetFcLcb(IndexPlcfFldHdr);
        (FcPlcfFldFtn, LcbPlcfFldFtn)   = GetFcLcb(IndexPlcfFldFtn);
        (FcPlcfFldAtn, LcbPlcfFldAtn)   = GetFcLcb(IndexPlcfFldAtn);
        (FcSttbfBkmk, LcbSttbfBkmk)     = GetFcLcb(IndexSttbfBkmk);
        (FcPlcfBkf, LcbPlcfBkf)         = GetFcLcb(IndexPlcfBkf);
        (FcPlcfBkl, LcbPlcfBkl)         = GetFcLcb(IndexPlcfBkl);
        (FcDop, LcbDop)                 = GetFcLcb(IndexDop);
        (FcClx, LcbClx)                 = GetFcLcb(IndexClx);
        (FcPlcSpaMom, LcbPlcSpaMom)     = GetFcLcb(IndexPlcSpaMom);
        (FcPlcfAtnbkf, LcbPlcfAtnbkf)   = GetFcLcb(IndexPlcfAtnbkf);
        (FcPlcfAtnbkl, LcbPlcfAtnbkl)   = GetFcLcb(IndexPlcfAtnbkl);
        (FcPlcfendRef, LcbPlcfendRef)   = GetFcLcb(IndexPlcfendRef);
        (FcPlcfendTxt, LcbPlcfendTxt)   = GetFcLcb(IndexPlcfendTxt);
        (FcPlcfFldEdn, LcbPlcfFldEdn)   = GetFcLcb(IndexPlcfFldEdn);
        (FcSttbfRgtlv, LcbSttbfRgtlv)   = GetFcLcb(IndexSttbfRgtlv);
        (FcTxbx, LcbTxbx)               = GetFcLcb(IndexTxbxText);
        (FcPlcfFldTxbx, LcbPlcfFldTxbx) = GetFcLcb(IndexPlcfFldTxbx);

        (FcFtn, LcbFtn)                 = (FcPlcffndTxt, LcbPlcffndTxt);
        (FcEnd, LcbEnd)                 = (FcPlcfendTxt, LcbPlcfendTxt);
        (FcAnot, LcbAnot)               = (FcPlcfandTxt, LcbPlcfandTxt);

        if (cbRgFcLcb > IndexPlcfLst)
        {
            (FcPlcfLst, LcbPlcfLst) = GetFcLcb(IndexPlcfLst);
        }

        if (cbRgFcLcb > IndexPlfLfo)
        {
            (FcPlfLfo, LcbPlfLfo) = GetFcLcb(IndexPlfLfo);
        }
    }

    public void SetDerivedFootnoteCharacterCount(int footnoteCharacterCount)
    {
        if (footnoteCharacterCount > 0 && CcpFtn == 0)
        {
            CcpFtn = footnoteCharacterCount;
        }
    }

    public (uint fc, uint lcb) GetFcLcb(int index)
    {
        if (index >= 0 && index < _rgFcLcb.Count) return _rgFcLcb[index];
        return (0, 0);
    }
}
