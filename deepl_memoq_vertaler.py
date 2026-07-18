#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeepL2TMX
=========

Leest een eentalige .docx in, segmenteert de tekst op zinsniveau, vertaalt elke
zin via de DeepL API en schrijft het resultaat weg als:

  * een tweetalige TMX  (translation memory, voor pre-translate in je CAT-tool)
  * een tweetalige XLIFF (generiek XLIFF 1.2, bilingueel document, importeerbaar
    in vrijwel elke CAT-tool)

Alleen standaardbibliotheken worden bovenaan geimporteerd, plus python-docx.
`deepl` en `tkinter` worden pas geladen wanneer ze nodig zijn, zodat de
kernfuncties los te testen/importeren zijn.

Afhankelijkheden om te installeren (in een terminal):
    pip install --upgrade deepl python-docx
    pip install --upgrade pysbd        # optioneel, betere zinssegmentatie

Starten:
    python deepl_memoq_vertaler.py
"""

from __future__ import annotations

import html
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

from docx import Document
from docx.document import Document as _DocxDocument
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

APP_NAME = "DeepL2TMX"
APP_VERSION = "1.0"
CONFIG_PATH = Path.home() / ".deepl_memoq_vertaler.json"
BATCH_SIZE = 40  # aantal segmenten per DeepL-verzoek

# UI-naam -> (DeepL source-code, DeepL target-code, TMX/XLIFF xml:lang-code)
LANGUAGES = {
    "Engels": ("EN", "EN-US", "en"),
    "Nederlands": ("NL", "NL", "nl"),
    "Duits": ("DE", "DE", "de"),
    "Frans": ("FR", "FR", "fr"),
    "Spaans": ("ES", "ES", "es"),
}


# ---------------------------------------------------------------------------
# 1. .docx inlezen (in documentvolgorde: alinea's + tabelcellen)
# ---------------------------------------------------------------------------
def _iter_block_items(parent):
    """Levert Paragraph- en Table-objecten in de volgorde waarin ze in het
    document staan (python-docx biedt dit niet standaard aan)."""
    if isinstance(parent, _DocxDocument):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Niet-ondersteund bovenliggend element")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _iter_table_texts(table):
    """Doorloopt de echte <w:tc>-elementen per rij. Horizontale merges (gridSpan)
    verschijnen daardoor eenmalig; verticale-merge-continuaties worden overgeslagen."""
    for row in table.rows:
        for tc in row._tr.tc_lst:
            tcPr = tc.tcPr
            if tcPr is not None:
                vmerge = tcPr.find(qn("w:vMerge"))
                if vmerge is not None and vmerge.get(qn("w:val")) in (None, "continue"):
                    continue  # voortzetting van een verticaal samengevoegde cel
            cell = _Cell(tc, table)
            for block in _iter_block_items(cell):
                if isinstance(block, Paragraph):
                    if block.text and block.text.strip():
                        yield block.text
                elif isinstance(block, Table):
                    yield from _iter_table_texts(block)


def _iter_paragraph_texts(container):
    for block in _iter_block_items(container):
        if isinstance(block, Paragraph):
            text = block.text
            if text and text.strip():
                yield text
        elif isinstance(block, Table):
            yield from _iter_table_texts(block)


def read_docx_paragraphs(path) -> list[str]:
    """Alle niet-lege alineateksten uit een .docx, in documentvolgorde."""
    doc = Document(str(path))
    return list(_iter_paragraph_texts(doc))


# ---------------------------------------------------------------------------
# 2. Zinssegmentatie
# ---------------------------------------------------------------------------
_ABBREV = (
    r"Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|vs|etc|Inc|Ltd|Co|Corp|Dept|Fig|No|Vol|"
    r"pp|approx|cf|al|Ph\.D|e\.g|i\.e|a\.m|p\.m|U\.S|U\.K|Rev|Gen|Sen|Gov|Hon"
)
_PROT_ABBREV = re.compile(r"\b(" + _ABBREV + r")\.", re.IGNORECASE)
_PROT_INITIAL = re.compile(r"\b([A-Z])\.")
_PROT_DECIMAL = re.compile(r"(\d)\.(\d)")
_BOUNDARY = re.compile(r'(?<=[.!?])["\')\]]*\s+(?=[A-Z0-9"\'(\[])')


def _fallback_split(text: str) -> list[str]:
    """Eenvoudige, redelijk robuuste Engelse zinssplitter (fallback zonder pysbd)."""
    tmp = text.replace("...", "\x01")
    tmp = _PROT_ABBREV.sub(lambda m: m.group(1) + "\x00", tmp)
    tmp = _PROT_INITIAL.sub(lambda m: m.group(1) + "\x00", tmp)
    tmp = _PROT_DECIMAL.sub(lambda m: m.group(1) + "\x00" + m.group(2), tmp)
    parts = _BOUNDARY.split(tmp)
    out = []
    for part in parts:
        s = part.replace("\x00", ".").replace("\x01", "...").strip()
        if s:
            out.append(s)
    return out


_pysbd_seg = None
_pysbd_tried = False


def _get_pysbd(lang_code: str):
    global _pysbd_seg, _pysbd_tried
    if _pysbd_tried:
        return _pysbd_seg
    _pysbd_tried = True
    try:
        import pysbd  # type: ignore

        _pysbd_seg = pysbd.Segmenter(language=lang_code, clean=False)
    except Exception:
        _pysbd_seg = None
    return _pysbd_seg


def segment_paragraph(text: str, lang_code: str = "en") -> list[str]:
    seg = _get_pysbd(lang_code)
    if seg is not None:
        try:
            return [s.strip() for s in seg.segment(text) if s.strip()]
        except Exception:
            pass
    return _fallback_split(text)


def segment_document(paragraphs, lang_code: str = "en") -> list[str]:
    segments = []
    for para in paragraphs:
        segments.extend(segment_paragraph(para, lang_code))
    return segments


# ---------------------------------------------------------------------------
# 3. Vertalen via DeepL
# ---------------------------------------------------------------------------
class DeepLClient:
    def __init__(self, api_key: str):
        import deepl  # pas hier importeren

        self._deepl = deepl
        self.translator = deepl.Translator(api_key.strip())

    def usage(self):
        return self.translator.get_usage()

    def translate(self, texts, source_lang, target_lang, formality="default",
                  progress=None):
        """Vertaalt een lijst zinnen 1-op-1 en behoudt de volgorde."""
        formality_arg = None
        if formality and formality != "default":
            # prefer_* degradeert netjes bij talen zonder formaliteit
            formality_arg = {"formeel": "prefer_more",
                             "informeel": "prefer_less"}.get(formality, formality)

        results = []
        total = len(texts)
        for start in range(0, total, BATCH_SIZE):
            batch = texts[start:start + BATCH_SIZE]
            res = self.translator.translate_text(
                batch,
                source_lang=source_lang,
                target_lang=target_lang,
                split_sentences="0",          # al voorgesegmenteerd
                preserve_formatting=True,
                formality=formality_arg,
            )
            if isinstance(res, list):
                results.extend(r.text for r in res)
            else:
                results.append(res.text)
            if progress:
                progress(min(start + len(batch), total), total)
        return results


# ---------------------------------------------------------------------------
# 4. Uitvoer schrijven: TMX en XLIFF
# ---------------------------------------------------------------------------
def _now_tmx() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def write_tmx(pairs, src_xml, tgt_xml, out_path):
    """pairs: lijst van (bron, doel). Schrijft TMX 1.4b."""
    stamp = _now_tmx()
    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<tmx version="1.4">',
        f'  <header creationtool="{xml_escape(APP_NAME)}" '
        f'creationtoolversion="{APP_VERSION}" segtype="sentence" '
        f'o-tmf="plaintext" adminlang="en" srclang="{src_xml}" '
        f'datatype="plaintext" creationdate="{stamp}"/>',
        "  <body>",
    ]
    for src, tgt in pairs:
        lines.append(f'    <tu creationdate="{stamp}">')
        lines.append(f'      <tuv xml:lang="{src_xml}"><seg>{xml_escape(src)}</seg></tuv>')
        lines.append(f'      <tuv xml:lang="{tgt_xml}"><seg>{xml_escape(tgt)}</seg></tuv>')
        lines.append("    </tu>")
    lines.append("  </body>")
    lines.append("</tmx>")
    Path(out_path).write_text("\n".join(lines), encoding="utf-8")


def write_xliff(pairs, src_xml, tgt_xml, original_name, out_path):
    """pairs: lijst van (bron, doel). Schrijft XLIFF 1.2 met ingevulde targets."""
    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<xliff version="1.2" xmlns="urn:oasis:names:tc:xliff:document:1.2">',
        f'  <file original="{xml_escape(original_name)}" '
        f'source-language="{src_xml}" target-language="{tgt_xml}" '
        f'datatype="plaintext">',
        "    <body>",
    ]
    for i, (src, tgt) in enumerate(pairs, start=1):
        lines.append(f'      <trans-unit id="{i}">')
        lines.append(f'        <source xml:lang="{src_xml}">{xml_escape(src)}</source>')
        lines.append(f'        <target xml:lang="{tgt_xml}">{xml_escape(tgt)}</target>')
        lines.append("      </trans-unit>")
    lines.append("    </body>")
    lines.append("  </file>")
    lines.append("</xliff>")
    Path(out_path).write_text("\n".join(lines), encoding="utf-8")


# ---------------------------------------------------------------------------
# 5. Volledige verwerking (los aanroepbaar, GUI-onafhankelijk)
# ---------------------------------------------------------------------------
def _fmt(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def compute_stats(segments):
    n_seg = len(segments)
    n_words = sum(len(re.findall(r"\S+", s)) for s in segments)
    n_chars = sum(len(s) for s in segments)
    n_chars_ns = sum(len(re.sub(r"\s+", "", s)) for s in segments)
    return n_seg, n_words, n_chars, n_chars_ns


def write_stats(out_path, docx_name, src_xml, tgt_xml, stats):
    n_seg, n_words, n_chars, n_chars_ns = stats
    lines = [
        "Analyse",
        "=======",
        f"Bestand:           {docx_name}",
        f"Talencombinatie:   {src_xml} -> {tgt_xml}",
        f"Datum:             {datetime.now():%Y-%m-%d %H:%M}",
        "",
        f"Segmenten:                    {_fmt(n_seg)}",
        f"Woorden (bron):               {_fmt(n_words)}",
        f"Tekens (bron, incl. spaties): {_fmt(n_chars)}",
        f"Tekens (bron, excl. spaties): {_fmt(n_chars_ns)}",
        "",
        "Woorden geteld als reeksen tekst gescheiden door spaties; dit kan licht",
        "afwijken van de telling in je CAT-tool.",
    ]
    Path(out_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def output_paths(docx_path, out_dir, src_name, tgt_name):
    """Geeft de doelpaden {'tmx','xliff','stats'}, allemaal in out_dir. Zowel de
    GUI (conflictcheck) als process_file gebruiken deze functie, zodat de paden
    altijd identiek zijn."""
    src_xml = LANGUAGES[src_name][2]
    tgt_xml = LANGUAGES[tgt_name][2]
    stem = Path(docx_path).stem
    tag = f"{src_xml}-{tgt_xml}"
    out_dir = Path(out_dir)
    return {
        "tmx": out_dir / f"{stem}_{tag}.tmx",
        "xliff": out_dir / f"{stem}_{tag}.xlf",
        "stats": out_dir / f"{stem}_analyse.txt",
    }


def process_file(docx_path, api_key, src_name, tgt_name, out_dir,
                 formality="default", make_tmx=True, make_xliff=True,
                 make_stats=True, skip=None, log=print, progress=None):
    src_deepl, _src_tgtcode, src_xml = LANGUAGES[src_name]
    _t_src, tgt_deepl, tgt_xml = LANGUAGES[tgt_name]

    docx_path = Path(docx_path)
    paths = output_paths(docx_path, out_dir, src_name, tgt_name)
    skip = {str(p) for p in (skip or [])}

    do_tmx = make_tmx and str(paths["tmx"]) not in skip
    do_xliff = make_xliff and str(paths["xliff"]) not in skip
    do_stats = make_stats and str(paths["stats"]) not in skip

    for flag_on, will_do, key in ((make_tmx, do_tmx, "tmx"),
                                  (make_xliff, do_xliff, "xliff"),
                                  (make_stats, do_stats, "stats")):
        if flag_on and not will_do:
            log(f"Overgeslagen (bestaat al): {paths[key].name}")

    log(f"Bestand inlezen: {docx_path.name}")
    paragraphs = read_docx_paragraphs(docx_path)
    log(f"  {len(paragraphs)} alinea's gevonden.")

    segments = segment_document(paragraphs, src_xml)
    log(f"  {len(segments)} segmenten na segmentatie.")
    if not segments:
        raise ValueError("Geen tekst gevonden om te vertalen.")

    stats = compute_stats(segments)
    written = []

    if do_tmx or do_xliff or do_stats:
        Path(out_dir).mkdir(parents=True, exist_ok=True)

    if do_tmx or do_xliff:
        log(f"  {_fmt(stats[1])} woorden / ~{_fmt(stats[2])} tekens naar DeepL.")
        client = DeepLClient(api_key)
        log(f"Vertalen ({src_deepl} -> {tgt_deepl})...")
        translations = client.translate(
            segments, src_deepl, tgt_deepl, formality=formality, progress=progress
        )
        if len(translations) != len(segments):
            raise RuntimeError(
                f"Aantal vertalingen ({len(translations)}) wijkt af van aantal "
                f"segmenten ({len(segments)})."
            )
        pairs = list(zip(segments, translations))

        if do_tmx:
            write_tmx(pairs, src_xml, tgt_xml, paths["tmx"])
            written.append(paths["tmx"])
            log(f"TMX geschreven -> {paths['tmx']}")

        if do_xliff:
            write_xliff(pairs, src_xml, tgt_xml, docx_path.name, paths["xliff"])
            written.append(paths["xliff"])
            log(f"XLIFF geschreven -> {paths['xliff']}")

    if do_stats:
        write_stats(paths["stats"], docx_path.name, src_xml, tgt_xml, stats)
        written.append(paths["stats"])
        log(f"Analyse geschreven -> {paths['stats']}")

    log("Klaar.")
    return written


# ---------------------------------------------------------------------------
# 6. Config
# ---------------------------------------------------------------------------
def load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(data: dict):
    try:
        CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 7. GUI
# ---------------------------------------------------------------------------
def launch_gui():
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    cfg = load_config()
    log_queue: "queue.Queue[tuple]" = queue.Queue()

    root = tk.Tk()
    root.title(f"{APP_NAME} {APP_VERSION}")
    root.geometry("640x560")
    root.minsize(560, 480)

    pad = {"padx": 10, "pady": 4}
    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    row = 0
    ttk.Label(frm, text="DeepL API-sleutel:").grid(row=row, column=0, sticky="w", **pad)
    key_var = tk.StringVar(value=cfg.get("api_key", ""))
    key_entry = ttk.Entry(frm, textvariable=key_var, show="*")
    key_entry.grid(row=row, column=1, columnspan=2, sticky="ew", **pad)
    remember_var = tk.BooleanVar(value=bool(cfg.get("api_key")))
    ttk.Checkbutton(frm, text="Onthouden", variable=remember_var).grid(
        row=row, column=3, sticky="w", **pad)

    row += 1
    ttk.Label(frm, text="Bronbestand (.docx):").grid(row=row, column=0, sticky="w", **pad)
    file_var = tk.StringVar(value="")

    def choose_file():
        init = ""
        if file_var.get():
            init = str(Path(file_var.get()).parent)
        elif cfg.get("last_src_dir"):
            init = cfg.get("last_src_dir")
        path = filedialog.askopenfilename(
            title="Kies een .docx-bestand",
            initialdir=init,
            filetypes=[("Word-document", "*.docx"), ("Alle bestanden", "*.*")],
        )
        if path:
            file_var.set(path)

    ttk.Entry(frm, textvariable=file_var).grid(row=row, column=1, columnspan=2, sticky="ew", **pad)
    ttk.Button(frm, text="Kies...", command=choose_file).grid(row=row, column=3, **pad)

    row += 1
    ttk.Label(frm, text="Van:").grid(row=row, column=0, sticky="w", **pad)
    src_var = tk.StringVar(value=cfg.get("src", "Engels"))
    ttk.Combobox(frm, textvariable=src_var, values=list(LANGUAGES), state="readonly",
                 width=15).grid(row=row, column=1, sticky="w", **pad)
    ttk.Label(frm, text="Naar:").grid(row=row, column=2, sticky="e", **pad)
    tgt_var = tk.StringVar(value=cfg.get("tgt", "Nederlands"))
    ttk.Combobox(frm, textvariable=tgt_var, values=list(LANGUAGES), state="readonly",
                 width=15).grid(row=row, column=3, sticky="w", **pad)

    row += 1
    ttk.Label(frm, text="Formaliteit:").grid(row=row, column=0, sticky="w", **pad)
    form_var = tk.StringVar(value=cfg.get("formality", "default"))
    ttk.Combobox(frm, textvariable=form_var,
                 values=["default", "formeel", "informeel"], state="readonly",
                 width=15).grid(row=row, column=1, sticky="w", **pad)

    row += 1
    tmx_var = tk.BooleanVar(value=cfg.get("tmx", True))
    xlf_var = tk.BooleanVar(value=cfg.get("xliff", True))
    ttk.Checkbutton(frm, text="TMX (voor je TM / pre-translate)", variable=tmx_var).grid(
        row=row, column=0, columnspan=2, sticky="w", **pad)
    ttk.Checkbutton(frm, text="XLIFF (bilingueel document)", variable=xlf_var).grid(
        row=row, column=2, columnspan=2, sticky="w", **pad)

    row += 1
    ttk.Label(frm, text="Uitvoermap:").grid(row=row, column=0, sticky="w", **pad)
    outdir_var = tk.StringVar(value=cfg.get("outdir", ""))
    outdir_entry = ttk.Entry(frm, textvariable=outdir_var)
    outdir_entry.grid(row=row, column=1, columnspan=2, sticky="ew", **pad)

    def choose_outdir():
        d = filedialog.askdirectory(title="Kies uitvoermap")
        if d:
            outdir_var.set(d)

    outdir_btn = ttk.Button(frm, text="Kies...", command=choose_outdir)
    outdir_btn.grid(row=row, column=3, **pad)

    row += 1
    stats_var = tk.BooleanVar(value=cfg.get("stats", True))
    ttk.Checkbutton(frm, text="Analyse schrijven (segmenten + woorden)",
                    variable=stats_var).grid(row=row, column=0, columnspan=3,
                                             sticky="w", **pad)

    row += 1
    run_btn = ttk.Button(frm, text="Vertalen")
    run_btn.grid(row=row, column=0, **pad)
    prog = ttk.Progressbar(frm, mode="determinate")
    prog.grid(row=row, column=1, columnspan=3, sticky="ew", **pad)

    row += 1
    frm.rowconfigure(row, weight=1)
    log_box = tk.Text(frm, height=12, wrap="word", state="disabled")
    log_box.grid(row=row, column=0, columnspan=4, sticky="nsew", **pad)

    def ui_log(msg):
        log_queue.put(("log", msg))

    def ui_progress(done, total):
        log_queue.put(("progress", (done, total)))

    def poll_queue():
        try:
            while True:
                kind, payload = log_queue.get_nowait()
                if kind == "log":
                    log_box.configure(state="normal")
                    log_box.insert("end", payload + "\n")
                    log_box.see("end")
                    log_box.configure(state="disabled")
                elif kind == "progress":
                    done, total = payload
                    prog.configure(maximum=total, value=done)
                elif kind == "done":
                    run_btn.configure(state="normal")
                    if payload:
                        messagebox.showinfo(
                            APP_NAME,
                            "Klaar. Geschreven bestanden:\n\n"
                            + "\n".join(Path(p).name for p in payload),
                        )
                elif kind == "error":
                    run_btn.configure(state="normal")
                    messagebox.showerror(APP_NAME, str(payload))
        except queue.Empty:
            pass
        root.after(100, poll_queue)

    def worker(path, key, src, tgt, out_dir, formality,
               mk_tmx, mk_xlf, mk_stats, skip):
        try:
            written = process_file(
                path, key, src, tgt, out_dir,
                formality=formality, make_tmx=mk_tmx, make_xliff=mk_xlf,
                make_stats=mk_stats, skip=skip, log=ui_log, progress=ui_progress,
            )
            log_queue.put(("done", written))
        except Exception as exc:  # noqa: BLE001
            ui_log(f"FOUT: {exc}")
            log_queue.put(("error", exc))

    def on_run():
        path = file_var.get().strip()
        key = key_var.get().strip()
        if (tmx_var.get() or xlf_var.get()) and not key:
            messagebox.showwarning(APP_NAME, "Vul je DeepL API-sleutel in.")
            return
        if not path or not Path(path).exists():
            messagebox.showwarning(APP_NAME, "Kies een geldig .docx-bestand.")
            return
        if not tmx_var.get() and not xlf_var.get() and not stats_var.get():
            messagebox.showwarning(APP_NAME,
                                   "Kies minstens een uitvoer (TMX, XLIFF of analyse).")
            return

        out_dir = outdir_var.get().strip()
        if not out_dir:
            messagebox.showwarning(APP_NAME, "Kies een uitvoermap.")
            return

        # Conflictcheck: per reeds bestaand doelbestand vragen om te overschrijven.
        planned = output_paths(path, out_dir, src_var.get(), tgt_var.get())
        selected = []
        if tmx_var.get():
            selected.append(planned["tmx"])
        if xlf_var.get():
            selected.append(planned["xliff"])
        if stats_var.get():
            selected.append(planned["stats"])

        skip = set()
        for target in selected:
            if target.exists():
                overwrite = messagebox.askyesno(
                    APP_NAME,
                    f"'{target.name}' bestaat al in:\n{target.parent}\n\nOverschrijven?",
                )
                if not overwrite:
                    skip.add(str(target))

        if selected and len(skip) == len(selected):
            messagebox.showinfo(
                APP_NAME, "Niets te doen: alle bestaande bestanden overgeslagen.")
            return

        save_config({
            "api_key": key if remember_var.get() else "",
            "src": src_var.get(),
            "tgt": tgt_var.get(),
            "formality": form_var.get(),
            "tmx": tmx_var.get(),
            "xliff": xlf_var.get(),
            "stats": stats_var.get(),
            "outdir": outdir_var.get(),
            "last_src_dir": str(Path(path).parent),
        })

        run_btn.configure(state="disabled")
        prog.configure(value=0)
        log_box.configure(state="normal")
        log_box.delete("1.0", "end")
        log_box.configure(state="disabled")

        threading.Thread(
            target=worker,
            args=(path, key, src_var.get(), tgt_var.get(), out_dir,
                  form_var.get(), tmx_var.get(), xlf_var.get(), stats_var.get(),
                  skip),
            daemon=True,
        ).start()

    run_btn.configure(command=on_run)
    root.after(100, poll_queue)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
