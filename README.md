# Local VN to JP Translation Agent

Local Vietnamese -> Japanese Markdown translation with
`translategemma:4b` through Ollama by default.

## Structure

```text
JP-VN-translation-agent/
├─ input/
├─ output/
├─ run_translate_all.py
├─ excel_translate.py
├─ glossary.json
├─ requirements.txt
├─ setup.bat
├─ run.bat
└─ configure_ollama_shared.bat
```

## First setup

1. Double-click `setup.bat`.
2. The script creates `.venv`, installs the Python Ollama client,
   checks the Ollama runtime/server, and downloads
   `translategemma:4b` if necessary.

The model download is large, so the first setup requires an Internet
connection and enough free disk space.

## Shared-machine Ollama configuration

If this machine also runs other local Ollama agents/models, run
`configure_ollama_shared.bat` once.

It sets:

```text
OLLAMA_MAX_LOADED_MODELS=1
OLLAMA_NUM_PARALLEL=1
```

Then fully quit Ollama from the Windows system tray and start it again.

The translator keeps TranslateGemma loaded while the batch is active, then
explicitly unloads it when the batch finishes or exits with an error. This
releases RAM for other local Ollama tools.

## Translate Markdown

1. Copy `.md` / `.markdown` files into `input`.
2. Double-click `run.bat`.
3. Results are written under `output`.

The translator recursively scans subdirectories and preserves their relative
directory structure.

Example:

```text
input/
├─ a.md
└─ design/
   ├─ b.md
   └─ api/
      └─ c.md
```

becomes:

```text
output/
├─ a.md
└─ design/
   ├─ b.md
   └─ api/
      └─ c.md
```

## Markdown protection

The translator protects structural/literal content such as:

- YAML front matter
- fenced code blocks
- inline code
- Markdown link destinations and raw URLs
- HTML/XML tags and comments
- Markdown table separator rows
- common template/variable expressions

Each chunk is validated after translation. Headings, list items, blockquotes,
table pipes, and protected placeholders must remain intact. Minor line-wrap
differences are allowed.

## Glossary

Edit `glossary.json` to enforce consistent Vietnamese -> Japanese terminology.

Example:

```json
{
  "màn hình": "画面",
  "chức năng": "機能"
}
```

If `glossary.json` is missing, translation still runs without a glossary.

## Command-line usage

```powershell
.\.venv\Scripts\python.exe run_translate_all.py
```

Custom directories:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --input-dir "D:\VI" --output-dir "D:\JA"
```

Custom glossary:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --glossary "D:\glossary.json"
```


## Performance modes

The default is the 4B Q4 model because it is much lighter on RAM and faster on
machines with limited memory:

```text
translategemma:4b
```

Default requests use at most about 1200 source characters per chunk and retry
at most twice. The terminal prints elapsed time, output token count, model load
time, and generation speed when Ollama reports those metrics.

For the slower 12B quality mode:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --model translategemma:12b-it-q4_K_M
```

You can also adjust chunk size:

```powershell
.\.venv\Scripts\python.exe run_translate_all.py --max-chars 800
```


## Excel and LibreOffice translation

Supported formats: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`.

The agent **automatically probes Microsoft Excel desktop**. If Excel COM is
available it uses that backend; if unavailable or opening fails it falls back
to the **LibreOffice UNO** backend. LibreOffice is sufficient for `.xlsx`
and `.xls`, and Microsoft Excel is not required. Microsoft Excel is still
recommended for files with strict layout fidelity, advanced drawings or macros.

On Windows, install LibreOffice normally (e.g.
`C:\Program Files\LibreOffice`); the agent detects `soffice.exe`. If
installed elsewhere, set `LIBREOFFICE_HOME` to the installation directory.
The UNO worker loads native Python bindings / DLLs from the installed
LibreOffice `program` directory and uses a dedicated headless process.
**Do not copy arbitrary LibreOffice DLLs into this repository**: they depend
on each other and must match the LibreOffice version and Python ABI. The
preferred dynamic dependency source is the installed LibreOffice runtime.

`setup.bat` installs the Ollama Python client and optionally pywin32 on
Windows; LibreOffice does not need to be installed with pip. For UNO the
agent probes the bundled LibreOffice Python or a compatible local Python
interpreter which can `import uno`.

### Workbook protection and fidelity

- Source workbooks are never modified.
- The output is written only after processing a file.
- Formula cells and numeric cells remain unchanged.
- Text cells, text boxes and standard drawing/group shapes are translated.
- Where UNO exposes formatted text portions, the agent attempts to preserve
  bold, color, italic and font size per portion.
- If translated text exceeds shape size, the agent estimates layout, tries
  shrinking to **at least 8pt**, then expands horizontally and then vertically.
- The UNO sizing check is approximate, not rendered overflow measurement.

**Limitations:** LibreOffice may re-serialize `.xlsx` formatting and drawing
objects, especially charts, threaded comments, SmartArt, embedded OLE,
conditional formatting and complex shapes. Excel native formatting fidelity
cannot be guaranteed by a LibreOffice round trip. For safety, `.xlsm`
(macro-enabled) and `.xlsb` **require Microsoft Excel**; the LibreOffice
fallback refuses them rather than risking macro loss or binary corruption.
Make a backup and visually review the output before delivering a critical
client workbook.

Run normally with `run.bat`; it automatically selects the backend and
preserves the relative input/output directory structure.


## Sparse scanning, context batching and checkpoints (LibreOffice)

For LibreOffice UNO, the agent queries nonempty text cells with Calc's
`queryContentCells(CellFlags.STRING)` instead of iterating a rectangular
`UsedRange`. The previous 150,000-position rectangle cap has been removed.
Formula and number cells are never sent to TranslateGemma.

- Cell and shape text is deduplicated via a translation cache: identical
  strings are translated once per workbook and reused for every occurrence.
- Different text strings are grouped into contextual batches of at most
  **16 strings / about 1100 characters**. Each is tagged with an ID; the
  model must return the matching IDs. Invalid responses fall back to
  separate strings, so text is not assigned to the wrong cell.
- **Every 75 completed cells/objects** (default), UNO saves a checkpoint
  workbook plus a JSON manifest including completed object IDs, cache and
  the original input file's SHA-256.
- On restart using the same input/output file names, checkpoints are
  validated and processing resumes at the next unfinished object.
- A checkpoint never overwrites the final output file; the final file is
  published only when the whole workbook finishes successfully.
- After successful completion, temporary checkpoint artifacts are deleted.

Intermediate files in the `output` folder are named
`<name>.checkpoint.xlsx` and `<name>.checkpoint.json` (or `.xls`).
Do not edit those files while translation is running or paused.

Rich-text portions remain separate so formatting can be applied per
translated run. The UNO shape overflow size calculation remains approximate.

**Both Microsoft Excel COM and LibreOffice UNO** now implement sparse text
scanning, global per-workbook translation cache, context batching, a checkpoint
after every 75 completed cell/shape items, and resume using input SHA-256.
Excel COM uses `SpecialCells` to find text constants; LibreOffice uses
`queryContentCells(CellFlags.STRING)`. Formula and numeric cells are not
translated. The two backends share `translation_batch.py` for ID-validated
contextual batching.

Checkpoint files use the same path convention, but the formats are backend
specific and should not be resumed across a backend switch. Set
`SPREADSHEET_BACKEND=excel` or `SPREADSHEET_BACKEND=libreoffice` to pin
a backend for a long-running/resumed job. If a backend changes, move aside
the old checkpoint pair first.

Excel COM continues to preserve native Excel formatting better than the
LibreOffice round trip. Neither backend has been integration-tested with every
shape type or the user's specific workbooks.


## Spreadsheet translation validation

Spreadsheet cell/shape translation uses a separate model mode from Markdown.
It **does not** reject translations because a Japanese result happens to
contain a numbered line (for example `1.`). Markdown structural validation
remains enabled only for `.md` and `.markdown` files.

Contextual batches still validate `[[SEG_NNN]]` identifiers to avoid
mapping a translation onto the wrong cell. When TranslateGemma 4B does not
preserve the identifiers, the affected strings are translated individually,
using spreadsheet text validation (non-empty output), instead of failing the
whole workbook. Existing Excel/LibreOffice checkpoint files and cached
translations remain reusable if the input file has not changed.


## Windows checkpoint file locks (LibreOffice)

The UNO backend uses **immutable, versioned checkpoint workbooks**. Each save
creates a new `<stem>.checkpoint.<unique-id>.xlsx` file (or `.xls`).
The `<stem>.checkpoint.json` manifest is atomically updated *after* the new
workbook has been completely saved and hashed. This avoids WinError 5 when
LibreOffice is still holding the previous checkpoint workbook open.

A previously created fixed-name `<stem>.checkpoint.xlsx` and its JSON
manifest are still accepted for resume. When the first new checkpoint is
saved successfully, the manifest starts pointing to versioned files.
Garbage collection skips locked old files rather than aborting translation;
any remaining stale checkpoint generation can be deleted once LibreOffice has
closed, after a successful run. **Do not delete the checkpoint JSON or its
referenced workbook before resuming a failed run.**

## Rich-text preservation and reviewer notes

Both backends keep strike-through text (deleted historical decisions must remain
visible), bold, colored segments, italics, and font sizes wherever the document
format exposes them. For UNO, original line breaks are reinserted as styled
runs instead of flattening the entire cell. When rich-text structure cannot be
recovered exactly, the translation stops rather than silently discarding
formatting. Shape shrinking respects the relative sizes of styled text runs
with a minimum of 8 pt. LibreOffice's visual overflow estimate is approximate.

On successful completion, each workbook has a matching UTF-8 CSV in output:

```text
<workbook-name>.shape_review.csv
```

The report identifies sheet, object ID/name, shrink/resize action, original
and resulting font size, width, and height. It includes a header even if
no shapes were changed. Review these objects visually before publishing.

Identity entries in `glossary.json`, such as `"API": "API"`,
`"WBS": "WBS"`, and `"ID": "ID"`, mean keep that exact term.
Japanese text containing Latin identifiers is skipped unless Vietnamese
characters requiring translation are present. Technical literals are
validated on model output and, if changed, translation is performed around
protected terms rather than silently accepting altered identifiers.

Contextual multi-segment data uses JSON strings to preserve literal
backslashes, quotes, and line breaks. If the model fails to return
complete ID-to-string mappings, the segment group falls back to safe,
individual translations. Run regression tests with:

```powershell
.\.venv\Scripts\python.exe -m unittest test_translation_batch.py -v
```

Checkpoints now incorporate the model name, glossary and batching/format
algorithm revision. **Checkpoints produced by previous versions lacking this
signature will not resume**: keep old files as a backup, but expect a full
new translation when running this version against an old checkpoint.


## Safe UNO rich-text recovery and manual review

UNO text-portion enumeration sometimes excludes paragraph separators or
normalizes line endings. The translator now aligns styled portions with the
complete text using a conservative character-level fallback. Mixed styles
(including strike-through, bold, italic, colors and different font sizes)
are retained rather than flattened into one style.

If a specific cell/shape has a non-whitespace mismatch that cannot be aligned
safely, **the original remains unchanged**. The cell/shape ID is recorded in
the resumable checkpoint, and the workbook continues processing. At successful
completion, the agent creates:

```text
<workbook-name>.rich_text_review.csv
```

This separate review CSV lists the sheet, cell/object ID, source text and
reason. Entries represent **items deliberately left untranslated** to prevent
format loss; they require manual review. An empty report contains only a
header. The existing `.shape_review.csv` continues to track shrunk/resized
objects. The existing checkpoint version and model/glossary signature remain
unchanged, so the current checkpoint can resume after this fix.

Run regression tests:

```powershell
.\.venv\Scripts\python.exe -m unittest test_rich_text_alignment.py -v
```


## Multi-language TranslateGemma from run.bat

\`run.bat\` remains the entry point. Double-clicking the file runs the default
**Vietnamese to Japanese** direction (\`vi -> ja\`). You can override it from
PowerShell or Command Prompt without calling Python manually:

\`\`\`powershell
.\run.bat --from ja --to vi
.\run.bat --from en --to ja
.\run.bat --from zh --to vi
.\run.bat --from vi --to ja --model translategemma:12b
.\run.bat --from ja --to en --glossary ".\glossaries\my-ja-en.json"
\`\`\`

You can still pass existing options such as \`--input-dir\`, \`--output-dir\`,
\`--model\`, \`--max-chars\` or \`--help\` directly to \`run.bat\`.

Pair glossaries are **directional**, not automatically reversed:

\`\`\`text
glossaries/
  vi-ja.json     # Existing technical glossary, including HTC and HTM
  ja-vi.json     # Optional Japanese -> Vietnamese glossary
  en-ja.json     # Optional English -> Japanese glossary
  ja-en.json     # Optional Japanese -> English glossary
\`\`\`

When a corresponding pair file does not exist, translation proceeds with
an empty glossary. \`--glossary PATH\` explicitly overrides the default; a
missing explicitly supplied file fails clearly rather than being ignored.
The original root \`glossary.json\` remains for backward compatibility or
manual use, but the new default is \`glossaries/vi-ja.json\`.

The translation prompt, source-language text selection in UNO/Excel, and
contextual batch prompts use \`--from\` and \`--to\`. The checkpoint signature
includes both language codes and the selected glossary, so changing direction
does not reuse translations from another language pair. Previous checkpoints
with older signatures may need a fresh run.

Language detection is heuristic, especially for mixed-language cells and
languages with overlapping scripts. Inspect the output, review CSVs and
technical terminology before use.


## Protect URLs, paths, and filenames from translation

Markdown, LibreOffice UNO, and Excel COM now identify technical literals
**before** sending text to TranslateGemma. These are never translated:

- Web URLs such as \`https://example.com/api?q=1\` or \`www.example.jp\`
- Windows paths like \`C:\work\docs\report.xlsx\` and UNC shares
- Unix and relative paths such as \`/opt/app/config.json\`,
  \`./reports/daily.xlsx\`, \`reports/daily.xlsx\`
- File names with extensions such as \`report.xlsx\`, \`config.json\`,
  \`資料_20261024.xlsx\`, and quoted filenames containing spaces

Only the surrounding ordinary text is translated. For spreadsheets, the
literal-safe splitter uses original text offsets across multiple rich-text
runs, preserving their individual styles (including strikethrough).
Technical literals are also excluded from the deduplication cache to
prevent a protected occurrence from suppressing translation elsewhere.
Markdown uses its existing validated placeholders to preserve literals.
The original file's name and output hierarchy remain unchanged.

Protection is based on recognized syntax, so ambiguous slash-separated
natural-language text can be treated as a path. A reviewer should check
unusual/custom identifier patterns. Checkpoint signatures have a new
format revision because the content transformation has changed.

Run the tests:

\`\`\`powershell
.\.venv\Scripts\python.exe -m unittest test_technical_literals.py -v
\`\`\`


## Automated XLSX integration tests

Run the test suite **without a real Ollama server/model or customer workbook**.
A local deterministic fake translator is used, but the spreadsheet backend,
UNO worker, rich-text serialization, checkpoint and review reports are real.

On Windows (LibreOffice installed):

\`\`\`powershell
git pull
.\test_integration.bat                  # fast 5-workbook smoke test
.\test_integration.bat --full           # all 20 synthetic workbooks
.\test_integration.bat --smoke --keep   # keep input/output XLSX for debugging
\`\`\`

You may also run \`python tests/run_integration.py --full\`. The script
automatically creates its own isolated LibreOffice profile and workbooks in
\`tests/artifacts\`, then reads the **saved output XLSX files back** using
LibreOffice UNO. It does not read your real \`input/\` or \`output/\` folders.

The 20 cases cover:

1. Basic translation
2. Contextual batching and duplicates
3. Already-Japanese cells skipped
4. Windows paths (literal backslashes)
5. URLs and query strings
6. Japanese filenames ending in \`.xlsx\`
7. Unix and relative paths
8. Windows UNC paths
9. Mixed bold/color/font-size rich text
10. Strikethrough as historical deleted content
11. Multiline italic/strikethrough rich text
12. Formulas and numeric cells
13. Merged cells and formatting
14. Text-box shape translation
15. Shape strikethrough and bold
16. Shape fit and CSV review
17. Multiple sheets and sparsely populated distant cells
18. Quoted filenames containing spaces
19. Japanese-to-Vietnamese translation
20. Injected failure after a 75-cell checkpoint, then resume

Reports are saved under \`tests/reports/<timestamp>_integration.json\`
with per-case PASS/FAIL and translation call counts. If any case fails,
the generated workbooks and worker logs are retained in \`tests/artifacts\`.
The process exits nonzero on failures, so this can also be wired into CI.

To explicitly test Microsoft Excel COM when desktop Excel is installed:

\`\`\`powershell
.\test_integration.bat --full --backend excel
\`\`\`

This is a separate optional run; it cannot pass on a LibreOffice-only PC.
Synthetic text checks do not establish linguistic translation quality.
For that, keep a separate manually reviewed Vietnamese/Japanese golden set.

**Status:** Fixture-generation and LibreOffice-reopen flow was exercised
with all 20 cases on a Linux LibreOffice installation. The full translation
end-to-end suite still needs to run on the intended Windows machine. Do not
interpret the existence of test code as evidence that all 20 translation
cases already pass.
