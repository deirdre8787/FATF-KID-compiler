# FATF-KID-compiler — Key Issue Document compiler

_(Imported from `key_issue_document_compiler_vf`.)_

FATF MER comment compilation tool. Ingests delegation responses to the
Simplified 5th Round ME comment template (.docx **or .pdf**) and produces an
analytical Excel workbook plus a FATF-styled compiled Word document.

## Two-phase flow (new)

Clicking **Check and Compile** first runs a full validation pass and shows a
popup listing everything that needs attention. Nothing is compiled until you
confirm.

The popup has two parts:

1. **Country assignment confirmation** – any file that does not declare a
   delegation inside the document ('Comments from' dropdown empty) is listed
   with the country guessed from its filename pre-filled. You confirm or
   correct it. Compilation is blocked until every such file has a country.
2. **Findings list** – colour-coded:
   - red **errors** (unreadable file, no comment rows, missing country)
   - amber **warnings** (dropdown country ≠ filename country, unknown
     jurisdiction name, empty Type/reference/comment cells, template
     placeholders left in, identical duplicate comments, two files from the
     same delegation)
   - blue **notices** (late file detected, PDF parsed with reduced
     formatting fidelity, empty Section auto-resolved to Recommendation N
     from a criteria reference)

## What changed in this revision

Fixed from live testing against the US MER submissions (China, Finland, Hong
Kong China, India, Israel, Korea, Singapore):

1. **Delegations were being renamed to the assessed country.** The Word builder
   substituted the skeleton's country name across the whole document *after*
   the comments were injected, so Singapore's comments came out attributed to
   the United States, and any comment text mentioning the skeleton country or a
   date was rewritten too. The shell is now scrubbed **before** the body is
   injected – delegation names and comment text are never touched.

2. **PDF submissions came out as a wall of ragged one-line paragraphs.** A PDF
   has no paragraphs, only printer line-wraps, and each wrapped line was
   becoming its own Word paragraph (41 paragraphs in one cell). PDF text is now
   un-wrapped back into real paragraphs, so a PDF submission reads identically
   to a Word one. Spacer columns, page-split rows and multi-line dropdown values
   ("Key\nRecommended\nActions") are handled too.

3. **Files whose template was altered read as zero comments.** Korea's file had
   the header row in one table and the comments in a second table (a sign of
   conversion or editing outside the template). A header-less table whose rows
   look like comment rows is now salvaged, and the file is flagged as altered
   for manual checking.

4. **A comment was being silently dropped.** India filed a comment with
   "Rec. 10" typed into the Section dropdown and no Type. Both read as
   unrecognised, so the row was discarded. Now: a Recommendation typed into the
   Section cell is understood, and – as a hard safety net – **a row carrying
   real comment text is never dropped**, whatever its dropdowns say. It is
   compiled under "Unspecified" and raised as a manual check.

5. **The KRA Roadmap table has an extra "IO" column** at the far left, showing
   the Immediate Outcome each Key Recommended Action was filed against. Its
   width is taken out of the Comment column, so the table is exactly as wide as
   every other table in the document. Only the KRA Roadmap has this column, and
   its rows are ordered by that IO (IO.1 → IO.11, then by delegation); rows with
   no IO sort to the end.

6. **KRA Roadmap comments now appear twice** – once in the KRA Roadmap section
   and again under the Immediate Outcome they were filed against. The KRA
   Roadmap section carries a note saying so. The Excel keeps one row per
   comment, so counts stay true to the source files.

7. **Nothing is typed by hand any more.** Every country – the assessed country
   and each delegation – is chosen from a dropdown of FATF delegations,
   observers/FSRBs and all other countries. Each file also has a Late / On time
   dropdown that overrides the filename rule. Aliases (USA, Republic of Korea,
   Turkey, Hong Kong SAR…) all resolve to one canonical name, so a delegation
   can never be split in two by spelling.

8. **A deleted delegation dropdown can no longer poison the output.** If the
   "Comments from" label survives but its dropdown has been removed, the next
   content control in the file is often the comment table itself – which came
   back as a "country" containing the whole table. Any detected jurisdiction
   that is not a plausible, recognisable delegation name is now rejected, the
   file is treated as "country not declared", and the filename guess is offered
   for confirmation instead.

9. **A file that yields zero comments hard-blocks the compile.** Fix and
   re-upload it, or tick "Accept empty" against it so the decision is recorded.

## The review step

Clicking **Check and Compile** parses everything and opens a review table –
one row per file showing what the tool detected, a country dropdown, a
timeliness dropdown, the comment count, and (where relevant) an "accept empty"
tick. Below it is the findings list. Nothing is compiled until you confirm.

## Manual Checks sheet

The diagnostic report opens with **Manual Checks** – the list of things the
tool cannot decide for you, each with what was found and what to do:

| Priority | Raised when |
|---|---|
| HIGH | file unreadable; template altered; country not declared; image inside a comment; comment that cannot be placed under any section; row rescued despite unreadable dropdowns |
| MEDIUM | PDF submission; a source row that was skipped |
| LOW | identical text from two delegations |


## Diagnostic report (new)

Every compilation now produces a third download, **`{Country}_Diagnostic_Report_{date}.xlsx`**,
and the same sheets are appended to the main analysis workbook so the
reconciliation always travels with the data.

| Sheet | What it answers |
|---|---|
| Diagnostic Summary | Does everything reconcile, in one number |
| Check by File | file name → country assigned → rows in source → skipped → duplicates removed → expected → actually in compilation → **Match YES/NO** |
| Check by Country | same reconciliation aggregated per delegation |
| Skipped Rows | every source row that did not become a comment, with the row number, its reference, and why |
| Duplicates Removed | what was removed, from which file/row, and which copy was kept |
| Same Text, Diff. Country | identical comment text from two or more delegations (kept – flagged for eyeballing) |
| Warnings | every error / warning / notice raised during validation |

The reconciliation arithmetic, per file and per country:

```
  rows in source table
− rows skipped (blank / placeholder / unreadable)
− duplicate rows removed
= comments expected  →  checked against comments actually compiled
```

Any mismatch is shown in red in the UI, in the report, and returned in the
API response.

## Deduplication

The same delegation submitting the same text against the same reference is
an accidental double entry: only the first copy is compiled, and every
removal is listed in the Duplicates Removed sheet. Identical text from
**different** delegations is kept (two countries making the same point is
normal) and reported separately for review.

## PDF and image warnings

- A delegation returning a **PDF** instead of the Word template raises a
  warning but is still processed; the warning notes that bold/italic/colour
  and bullets cannot be recovered from a PDF.
- **Embedded images and objects** (screenshots, pasted charts, OLE objects)
  are detected in both .docx and .pdf inputs. For Word files the exact
  comment rows containing them are named, e.g. *"2 embedded image/object(s)
  found, 2 of them inside comment cells (row 3 (Paragraph 274), row 6 (Table
  4.5))"*. Images are **not** carried into the compiled Word document, so
  these rows need adding by hand afterwards. Images outside the comment
  table (letterheads, logos) are reported as a notice only.

## Checks performed

Per file: readable? contains the 4-column comment table? declares a
delegation? filename country matches declared country? flagged late
(standalone token `late` in the filename)? Per row: Section, Type,
reference, comment all present? placeholder text left in? Cross-file:
identical (delegation, reference, text) duplicate comments; multiple files
from the same delegation.

## PDF ingestion

PDF templates are parsed with pdfplumber. Tables split across pages are
stitched, placeholders skipped, and criteria references (R.5, C.1.6, c2.1…)
bucket into Recommendation N exactly like the docx path. Rich text
(bold/italic/colour) inside PDF comments is not preserved – flagged as a
notice.

## Output metadata (Word)

- author / last modified by: `deirdre.luo@fatf-gafi.org` – nothing else
- created / modified timestamps: the moment of generation
- visible dates in the document: replaced with the generation day
- document cote (`FATF/ME(yyyy)N`): removed everywhere – body sentences,
  headers, footers, and `customXml/item5.xml`
- meeting / SharePoint metadata (`customXml/item2.xml`): emptied
- OECD content-type schema (`customXml/item4.xml`): emptied
- `docProps/app.xml`: Company blank, Template = Normal.dotm, word /
  character / page counts zeroed, original document titles removed
- `docProps/custom.xml`: empty

## Formatting

- fixed 3-column tables everywhere: 2.50 cm | 18.80 cm | 2.07 cm,
  `tblLayout=fixed`, identical across all IOs/Recommendations
- col 1 + col 3 left-aligned, col 2 justified
- header row light blue, Substantive/Editorial divider rows light grey
- late comments: whole row light pastel pink (#FFD9E2) in Word and in the
  Excel Compiled sheet; late detection from the filename token `late`
  (`_late`, `late.`, `-late-` count; `later`, `slate`, `plate` do not)

## Run

Windows: double-click `run.bat`. Manual:

```bash
python -m venv .venv
source .venv/bin/activate           # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5050.

## Layout

```
FATF-KID-compiler/
├── app.py                # Flask server: /validate + /process two-phase flow
├── extractor.py          # DOCX → structured Comment records
├── pdf_extractor.py      # PDF fallback ingestion (pdfplumber)
├── validator.py          # pre-compilation checks / issue list
├── countries.py          # canonical names, aliases, dropdown lists
├── media_probe.py        # embedded image / object detection
├── diagnostics.py        # dedupe + source-to-compilation reconciliation
├── diag_export.py        # diagnostic report workbook
├── excel_export.py       # openpyxl workbook builder
├── word_export.py        # FATF-styled Word builder + metadata wipe
├── sorter.py             # bucket ordering
├── fatf_skeleton.docx    # bundled FATF skeleton
├── templates/index.html  # UI with validation popup
├── requirements.txt
├── run.bat
└── README.md
```
