# deepl2memoq

A small local Python app that turns a monolingual `.docx` into DeepL-translated,
sentence-aligned **TMX** and **XLIFF** files ready for memoQ.

It reads a Word document, splits the text into sentence-level segments, translates
them via the DeepL API, and writes an aligned bilingual TMX (for translation memory
import) and XLIFF (for direct editing in the grid), plus a plain-text analysis of
segment, word and character counts.

## Requirements

- Python 3.9+
- A DeepL API key (Free, Pro or Growth – the app selects the right endpoint automatically)
- Packages: see `requirements.txt`

```
pip install -r requirements.txt
```

`pysbd` is optional but improves sentence segmentation; without it a built-in
fallback splitter is used.

## Usage

```
python deepl_memoq_vertaler.py
```

A window opens. Enter your DeepL API key, pick the source `.docx`, choose the
languages, and click **Vertalen**. The key can be remembered; it is then stored in
`~/.deepl_memoq_vertaler.json` in your home folder – never inside this project.

## Output routing

By default the app follows a fixed project structure and creates missing folders:

```
Project folder/
├── Translatables/     <- source .docx is picked from here
├── Reference files/   <- TMX + XLIFF are written here
└── Statistics/        <- analysis (.txt) is written here
```

The project folder is derived from the source file (the folder above
`Translatables`). A manual mode lets you send all output to a folder of your choice.

Existing files are never overwritten without a per-file confirmation. If both TMX
and XLIFF are skipped, DeepL is not called at all (no quota is used).

## Two memoQ workflows

- **XLIFF** – import as a *document*; the translation lands directly in the grid and
  your segmentation is preserved. Most reliable, but the original layout is dropped
  (plain text).
- **TMX** – import as a *translation memory*, then import your original `.docx` as
  the document and run Pre-translate. Keeps the document's formatting, but exact
  matches depend on memoQ's segmentation matching the app's.

## Notes

- Inline formatting and tags (bold, italics, etc.) are flattened to plain text.
- Word counts are approximate (whitespace-separated tokens) and may differ slightly
  from memoQ's count.
