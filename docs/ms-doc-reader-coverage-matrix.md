# MS-DOC Reader Coverage Matrix

This matrix tracks the read layer against the major MS-DOC structures exposed by FibReader. It is meant to distinguish three states:

- Covered: parsed structurally and fed into the document model.
- Partial: parsed with best-effort, heuristics, or lossy flattening.
- Missing: exposed by FIB but not meaningfully consumed.

## Core Structures

| Structure | FIB Entry | Reader Path | Status | Notes |
| --- | --- | --- | --- | --- |
| CFB container | streams | Readers/CfbReader.cs | Covered | WordDocument, Table, Data, optional story streams are opened. |
| FIB base and fc/lcb map | WordDocument prefix | Readers/FibReader.cs | Covered | Primary routing table for all downstream readers. |
| CLX and piece table | FcClx/LcbClx | Readers/TextReader.cs | Covered | Main text reconstruction and CP/FC mapping are implemented. |
| CHP/PAP FKPs | FcPlcfBteChpx/FcPlcfBtePapx | Readers/FkpParser.cs | Covered | Used for body paragraph/run formatting. |
| SPRM application | grpprl payloads | Readers/SprmParser.cs | Partial | Broad opcode coverage exists, but still grows through regressions and samples. |
| STSH styles | FcStshf/LcbStshf | Readers/StyleReader.cs | Partial | Real STD parsing exists, but defaults and fallbacks still mask unsupported variants. |
| Fonts | FcSttbfFfn/LcbSttbfFfn | Readers/StyleReader.cs | Partial | FFN parsing exists, malformed ranges fall back to defaults. |
| DOP | FcDop/LcbDop | Readers/DocReader.cs plus property reader | Partial | Core document properties are consumed; many DOP flags are not validated end to end. |

## Lists and Numbering

| Structure | FIB Entry | Reader Path | Status | Notes |
| --- | --- | --- | --- | --- |
| LST/LVLF | FcPlcfLst/LcbPlcfLst | Readers/ListReader.cs | Partial | Base list definitions are parsed; complex numbering variants still need stronger spec coverage. |
| LFO / list instances | FcPlfLfo/LcbPlfLfo | Readers/ListReader.cs | Partial | LFO headers plus LFOLVL start-at and numbering-format overrides are parsed structurally and now flow into numbering.xml; uncommon LFOLVL payload variants still need broader specimen coverage. |
| Paragraph ilfo/ilvl | SPRM opcodes | Readers/SprmParser.cs + Readers/DocReader.cs | Covered | Paragraph list instance ids flow into the model and numbering writer. |

## Stories

| Structure | FIB Entry | Reader Path | Status | Notes |
| --- | --- | --- | --- | --- |
| Main story | CcpText | Readers/DocReader.cs | Covered | Full paragraph/run parsing path. |
| Footnote story | FcFtn/LcbFtn plus CcpFtn | Readers/FootnoteReader.cs + Readers/DocReader.cs | Partial | Story ranges are reparsed structurally and field models are now extracted, but dedicated reader logic still starts from simplified note extraction. |
| Endnote story | FcEnd/LcbEnd plus CcpEdn | Readers/FootnoteReader.cs + Readers/DocReader.cs | Partial | Same caveat as footnotes; structured field models are now extracted from reparsed story content. |
| Annotation story | FcAnot/LcbAnot via Plcfs | Readers/AnnotationReader.cs + Readers/DocReader.cs | Partial | Ranges/authors are parsed and DocReader now reparses story content into paragraphs/runs/fields, but annotation metadata remains basic. |
| Textbox story | FcTxbx/LcbTxbx plus CcpTxbx | Readers/TextboxReader.cs + Readers/DocReader.cs | Partial | Story bounds are parsed, reparsed structurally, and now emit field models; anchoring and shape association still rely on heuristics. |
| Header/footer story | FcPlcfHdd/LcbPlcfHdd plus CcpHdd | Readers/HeaderFooterReader.cs + Readers/DocReader.cs | Partial | PLC parsing exists and DocReader reparses content into paragraphs plus field models, but matching and filtering still use heuristics. |
| Header textbox story | CcpHdrTxbx | no dedicated reader | Missing | Count is included in global text reconstruction, but no separate structured reader exists. |

## PLC Families and Ancillary Data

| Structure | FIB Entry | Reader Path | Status | Notes |
| --- | --- | --- | --- | --- |
| Bookmarks | FcPlcfBkf/FcPlcfBkl/FcSttbfBkmk | Readers/BookmarkReader.cs | Partial | Core bookmark ranges and names are parsed. Complex bookmark variants are not fully validated. |
| Fields in main story | FcPlcfFldMom | Readers/DocReader.cs + Readers/FieldReader.cs | Partial | Body field characters are interpreted during run parsing. |
| Fields in headers/footnotes/annotations/endnotes | FcPlcfFldHdr/Ftn/Atn/Edn | Readers/DocReader.cs | Partial | Field PLCs are read and validated against story text for diagnostics, but no dedicated field-model reader exists yet. |
| Fields in textbox story | FcPlcfFldTxbx | Readers/DocReader.cs | Partial | Used only for textbox anchor hints, not as a full story field parser. |
| Sections | FcPlcfSed/LcbPlcfSed | Readers/SectionReader.cs | Covered | SEPX records are applied with defaults on failure. |
| Floating shape anchors | FcPlcSpaMom/LcbPlcSpaMom | Readers/FspaReader.cs | Partial | Bounding boxes and anchors are best-effort. |
| Revision authors | FcSttbfRgtlv/LcbSttbfRgtlv | Readers/DocReader.cs | Covered | String table is loaded into the model. |

## Binary Objects and Drawing

| Structure | FIB Entry | Reader Path | Status | Notes |
| --- | --- | --- | --- | --- |
| OfficeArt / Escher | Data stream and drawing records | Readers/OfficeArtMapper.cs | Partial | Shapes and anchors are best-effort; advanced drawing semantics are incomplete. |
| Images | picf and Data stream references | Readers/DocReader.cs | Partial | Many cases work, but image extraction still includes best-effort branches. |
| Embedded OLE objects | ObjectPool and storages | Readers/DocReader.cs | Partial | Detection and chart recovery remain heuristic-heavy. |
| Charts | embedded BIFF | Readers/BiffChartScanner.cs | Partial | Minimal chart data recovery only. |
| VBA | macros storage | Readers/DocReader.cs | Partial | Extraction exists, but no semantic validation. |

## Priority Follow-Ups

1. Replace the remaining simplified LFO and LFOLVL parsing with spec-structured instance and level override decoding.
2. Add dedicated field PLC readers for header, footnote, annotation, endnote, and textbox stories.
3. Add a structured reader for header textbox stories instead of only counting them through global CcpHdrTxbx.
4. Reduce best-effort branches in OfficeArt, OLE, and image extraction by converting current warnings into explicit unsupported-case diagnostics.