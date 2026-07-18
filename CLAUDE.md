# CLAUDE.md

Guidance for Claude Code when working in this repository.

## What this project is

`deepl2memoq` is a single-file desktop tool. It reads a monolingual `.docx`,
segments the text into sentences, translates each segment via the DeepL API, and
writes an aligned bilingual **TMX** and **XLIFF** (plus a plain-text analysis) for
use in memoQ. All logic lives in `deepl_memoq_vertaler.py`.

## Run & develop

```
pip install -r requirements.txt      # deepl, python-docx, pysbd (pysbd optional)
python deepl_memoq_vertaler.py        # launches the tkinter window
```

There is no build step and no test suite yet.

## File layout of `deepl_memoq_vertaler.py` (in order)

1. Constants: `LANGUAGES`, `BATCH_SIZE`, `CONFIG_PATH`.
2. `.docx` reading: `_iter_block_items`, `_iter_table_texts`, `read_docx_paragraphs`.
3. Segmentation: `_fallback_split`, pysbd loader, `segment_paragraph`, `segment_document`.
4. `DeepLClient` — thin wrapper around the `deepl` library.
5. Output writers: `write_tmx`, `write_xliff`.
6. Analysis + routing: `compute_stats`, `write_stats`, `output_paths`.
7. `process_file` — orchestration (read → segment → translate → write).
8. Config: `load_config`, `save_config`.
9. `launch_gui` — the tkinter interface.

## Key design decisions (don't undo these without reason)

- **Lazy imports.** `deepl` and `tkinter` are imported *inside* the functions that
  use them, never at module top level. This keeps the core functions importable and
  testable in an environment without those packages. Keep it that way.
- **Segment-level 1:1 alignment.** Sentences are sent to DeepL as a list; DeepL
  returns results in the same order, giving clean source↔target pairs for TMX/XLIFF.
  `split_sentences="0"` is set because the text is already segmented. Never merge or
  reorder segments between source and translation.
- **Output routing.** The user always picks the output folder by hand; TMX, XLIFF
  and the analysis all go into that one folder. There is no automatic project-folder
  detection. `output_paths` is the single source of truth for filenames — both the
  GUI conflict check and `process_file` call it, so they must stay in sync.
- **Overwrite protection.** The GUI checks for existing target files *before*
  starting the worker thread and asks per file (a dialog may only run on the main
  thread). Skipped paths are passed to `process_file` as `skip`. If both TMX and
  XLIFF end up skipped, DeepL is not called at all (saves quota). Preserve this
  main-thread-first pattern for any new confirmation prompts.
- **Threading.** Translation runs in a background thread; it talks to the GUI only
  through a `queue.Queue` polled by `root.after`. Do not touch tkinter widgets from
  the worker thread.

## Gotchas

- **Table reading.** `_iter_table_texts` iterates the real `<w:tc>` elements per row
  and skips vertical-merge continuations. An earlier version deduplicated cells with
  `id(cell._tc)`, which is **buggy** — lxml hands out fresh Python wrappers whose
  `id()` collides after garbage collection, silently dropping cells. Do not
  reintroduce `id()`-based dedup.
- **DeepL Free vs Pro/Growth.** The `deepl` library auto-selects the endpoint from
  the key (only Free keys end in `:fx`). No manual endpoint switching needed.
- **Plain text only.** Inline formatting/tags are flattened; word counts are
  whitespace-based approximations and may differ slightly from memoQ.

## Testing convention

The core was verified without network or DeepL access by (a) generating a sample
`.docx` with python-docx, (b) shimming a fake `deepl` module on `PYTHONPATH`, and
(c) checking that TMX/XLIFF parse as valid XML and that segment/translation counts
line up. Keep new logic in the non-GUI functions so it can be tested this way; keep
`process_file` free of tkinter references.

## Conventions

- **Language split:** UI strings and code comments are in **Dutch**; the README and
  this file are in **English**. Keep that split when editing.
- In any prose/docs, use **en-dashes**, not em-dashes.
- Keep everything in the single script unless there's a clear reason to split it.

## Never

- Never commit or log a DeepL API key. The key lives in
  `~/.deepl_memoq_vertaler.json` (home folder, outside the repo) and is git-ignored.
