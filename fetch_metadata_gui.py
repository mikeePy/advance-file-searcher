"""
fetch_metadata_gui.py
=====================
Desktop GUI for fetching movie posters and NFO metadata from TMDB.

Key features in this version
----------------------------
• Strong query cleaner
    - strips [YTS.MX] / 【...】 / {...} blocks
    - strips (720p...) / (Extended) / (Directors.Cut) parentheses
    - strips release groups: YTS, YIFY, VXT, AcEs, Ganool, CMCT, BYNDR,
      Flux, GalaxyRG, ETRG, Fugi, Deceit, BOKUTOX, P2P, etc.
    - strips quality / codec / language tokens: 720p, BluRay, WEB-DL,
      HDRip, x264, x265, AAC, DDP, Atmos, Extended, JAPANESE, subs, …
    - strips episode markers: "Episode 1", "S01E02"
    - extracts a year (1900-2099) as the LAST match found
    - keeps Unicode titles (Chinese, Japanese, Korean) intact

• Query variants (fixes Ant-Man, Kiki's Delivery Service, etc.)
    - "Ant Man"  (hyphen → space)
    - "AntMan"   (hyphen removed)
    - "Ant-Man"  (original)
    - "Kikis Delivery Service" (apostrophes removed)

• Search resilience
    - tries every variant WITH year first, then WITHOUT year
    - handles HTTP 429 (rate-limit) gracefully
    - logs every attempt so you can see what was tried

• Editable query per row
    - right-click → "✏️ Edit query…"
    - or double-click the Query / Year column
    - after editing, the row clears its old match and can be re-searched

Requirements:
    pip install requests

Run:
    python fetch_metadata_gui.py
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import threading
import time
import queue
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    import requests
except ImportError:
    _msg = ("This app requires the 'requests' library.\n\n"
            "Install it with:\n\n    pip install requests")
    try:
        import tkinter as tk
        from tkinter import messagebox
        _r = tk.Tk(); _r.withdraw()
        messagebox.showerror("Missing dependency", _msg)
        _r.destroy()
    except Exception:
        print(_msg)
    sys.exit(1)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------

CONFIG_LOCATIONS = [
    Path.home() / ".tmdb_fetch_config.json",
    Path(".tmdb_fetch_config.json"),
]


def load_config() -> dict:
    for p in CONFIG_LOCATIONS:
        try:
            if p.exists():
                return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
    return {}


def save_config(data: dict) -> None:
    for p in CONFIG_LOCATIONS:
        try:
            p.write_text(json.dumps(data, indent=2), encoding="utf-8")
            return
        except Exception:
            continue


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

LANGUAGES = [
    "en-US", "zh-CN", "zh-TW", "zh-HK", "ja-JP", "ko-KR",
    "fr-FR", "de-DE", "es-ES", "it-IT", "pt-PT", "ru-RU",
]
IMAGE_SIZES = ["w342", "w500", "w780", "original"]

TMDB_SEARCH_URL = "https://api.themoviedb.org/3/search/movie"
TMDB_MOVIE_URL = "https://api.themoviedb.org/3/movie/{id}"
TMDB_IMAGE_URL = "https://image.tmdb.org/t/p/{size}{path}"

# Release-group suffixes — deleted when they appear as standalone tokens
RELEASE_GROUPS = [
    r"YTS(?:\.MX|\.AG|\.AM|\.PE|\.LT)?",
    r"YIFY",
    r"VXT(?:-AcEs)?",
    r"AcEs",
    r"Ganool(?:\.?iconvid\.co)?",
    r"CMCT",
    r"BYNDR",
    r"Flux",
    r"GalaxyRG",
    r"Tv21\.org",
    r"iconvid\.co",
    r"Vegamovies(?:\.is|\.hot|\.com)?",
    r"scOrp[_\s]*sujaidr",
    r"BOKUTOX",
    r"Deceit",
    r"ETRG",
    r"P2P",
    r"Fugi",
    r"ZhuixinFan(?:V2)?",
    r"ENG[\.\s]*HC",
    r"ENGHC",
    r"AEF",
    r"x264-[A-Za-z0-9]+",
    r"x265-[A-Za-z0-9]+",
]

# Quality / codec / language / metadata tokens — deleted as whole words
NOISE_TOKENS = (
    r"720p|1080p|2160p|480p|360p|4k|uhd|hd|"
    r"bluray|blu-ray|brrip|bdrip|webrip|web-dl|webdl|web|"
    r"hdrip|dvdrip|hdtv|dvd|"
    r"x264|x265|h\.?264|h\.?265|hevc|"
    r"aac|ac3|dts|ddp\d?|dd\d|atmos|"
    r"10bit|8bit|5\.1|7\.1|2\.0|"
    r"repack|proper|extended|remastered|directors?\.?cut|"
    r"hc|eng|engsub|engsubs|subs|subbed|subtitles|"
    r"multi|multisub|japanese|chinese|jap|"
    r"hin-jap|hindi|english|dual\.?audio|dubbed|"
    r"cineby|kisskh|mkv|mp4|avi|"
    r"vol\.?\s*\d+|disc\s*\d+"
)


# ---------------------------------------------------------------------------
# Query cleaner
# ---------------------------------------------------------------------------


def extract_years(text: str) -> list[str]:
    return re.findall(r"\b(19\d{2}|20\d{2})\b", text)


def clean_and_extract(name: str) -> dict:
    """
    Turn a messy folder name into {title, year, variants}.

    "Ant.Man.2015.720p.HDRip.x264.AAC-ETRG"
        → title="Ant Man", year="2015",
          variants=["Ant Man", "AntMan", ...]

    "The.Lord.of.the.Rings.The.Two.Towers.EXTENDED.2002.720p…"
        → title="The Lord of the Rings The Two Towers", year="2002"

    "滿江紅"
        → title="滿江紅", year=None, variants=["滿江紅"]
    """
    s = name

    # 1. Remove bracketed segments entirely
    s = re.sub(r"\[[^\]]*\]", " ", s)
    s = re.sub(r"【[^】]*】", " ", s)
    s = re.sub(r"\{[^}]*\}", " ", s)

    # 2. Remove parenthesised segments that contain quality markers
    s = re.sub(
        r"\([^)]*(?:YTS|BluRay|WEB|HDRip|BRRip|DVDrip|720p|1080p|480p|"
        r"2160p|4K|UHD|AAC|x264|x265|H264|H265|DDP|HDTV|Remastered|"
        r"Extended|Directors?\.?\s*Cut)[^)]*\)",
        " ", s, flags=re.I,
    )

    # 3. Remove leading junk like "720p HD MOVIE "
    s = re.sub(r"^(?:720p|1080p|2160p|4k)\s+(?:hd\s+)?movie\s+",
               "", s, flags=re.I)

    # 4. Replace separators
    s = s.replace(".", " ").replace("_", " ")

    # 5. Release-group suffixes
    for grp in RELEASE_GROUPS:
        s = re.sub(grp, " ", s, flags=re.I)

    # 6. Trailing "-GROUPNAME" fragments
    s = re.sub(r"-\s*[A-Za-z0-9_]{2,15}\s*$", " ", s)

    # 7. Noise tokens (word boundaries so "HD" won't match inside words)
    s = re.sub(r"\b(?:" + NOISE_TOKENS + r")\b", " ", s, flags=re.I)

    # 8. Episode markers
    s = re.sub(r"\bepisode\s+\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\bep\.?\s*\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\bs\d{1,2}\s*e\d{1,2}\b", " ", s, flags=re.I)

    # 9. Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip(" -.:;,·|/")

    # 10. Year extraction — take the LAST match found
    years = extract_years(s)
    year = years[-1] if years else None

    # 11. Remove ONLY the last year occurrence from the title
    if year:
        idx = s.rfind(year)
        title = (s[:idx] + s[idx + len(year):]).strip(" -.:;,·|/")
        if not title:
            title = year
    else:
        title = s

    # 12. Build search variants
    variants: list[str] = []
    if title:
        variants.append(title)                                   # original
        variants.append(re.sub(r"\s+", " ", title.replace("-", " ")).strip())
        variants.append(title.replace("-", ""))
        variants.append(title.replace("'", "").replace("’", ""))
        variants.append(title.replace(":", ""))
        variants.append(re.sub(r"\s+", " ", title.replace("-", " ")
                                .replace("'", "").replace("’", "")).strip())

    # Deduplicate, preserving order
    seen: set[str] = set()
    unique: list[str] = []
    for v in variants:
        v = v.strip()
        if v and v not in seen:
            seen.add(v)
            unique.append(v)

    return {"title": title, "year": year, "variants": unique}


def build_variants(title: str) -> list[str]:
    """Variants derived from an already-cleaned title."""
    if not title:
        return []
    candidates = [
        title,
        re.sub(r"\s+", " ", title.replace("-", " ")).strip(),
        title.replace("-", ""),
        title.replace("'", "").replace("’", ""),
        title.replace(":", ""),
        re.sub(r"\s+", " ", title.replace("-", " ")
               .replace("'", "").replace("’", "")).strip(),
    ]
    seen, out = set(), []
    for c in candidates:
        c = c.strip()
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


# ---------------------------------------------------------------------------
# TMDB calls
# ---------------------------------------------------------------------------


def search_movie_robust(
    title: str,
    year: str | None,
    api_key: str,
    language: str,
    include_adult: bool,
    log_cb=None,
):
    """
    Try every title variant, first with year, then without.
    Returns (hit_dict, variant_used, year_used) or (None, None, None).
    """
    variants = build_variants(title)
    if not variants:
        return None, None, None

    attempts: list[tuple[str, str | None]] = []
    for v in variants:
        if year:
            attempts.append((v, year))
        attempts.append((v, None))

    for v, y in attempts:
        params = {
            "api_key": api_key,
            "query": v,
            "language": language,
            "include_adult": "true" if include_adult else "false",
            "page": 1,
        }
        if y:
            params["year"] = y
        try:
            r = requests.get(TMDB_SEARCH_URL, params=params, timeout=15)
            if r.status_code == 429:
                if log_cb:
                    log_cb("  ! rate-limited, sleeping 2 s")
                time.sleep(2)
                continue
            r.raise_for_status()
            results = r.json().get("results", [])
            if results:
                if log_cb:
                    log_cb(f"  ✓ query={v!r} year={y!r} → "
                           f"{results[0].get('title')}")
                return results[0], v, y
            else:
                if log_cb:
                    log_cb(f"  · query={v!r} year={y!r} → 0 results")
        except Exception as exc:
            if log_cb:
                log_cb(f"  ! query={v!r} year={y!r}: {exc}")
        time.sleep(0.1)

    return None, None, None


def tmdb_details(movie_id: int, api_key: str, language: str) -> dict:
    params = {
        "api_key": api_key,
        "language": language,
        "append_to_response": "credits",
    }
    r = requests.get(
        TMDB_MOVIE_URL.format(id=movie_id), params=params, timeout=15
    )
    r.raise_for_status()
    return r.json()


def download_poster(poster_path: str, dest: Path, image_size: str) -> None:
    url = TMDB_IMAGE_URL.format(size=image_size, path=poster_path)
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    if len(r.content) < 1000:
        raise ValueError(f"Poster too small ({len(r.content)} bytes)")
    dest.write_bytes(r.content)


def write_nfo(movie: dict, folder: Path, actors_count: int) -> Path:
    """Write a Kodi-compatible NFO named after the first .mp4 (or folder)."""
    mp4_files = sorted(folder.glob("*.mp4"))
    nfo_stem = mp4_files[0].stem if mp4_files else folder.name
    nfo_path = folder / f"{nfo_stem}.nfo"

    root = ET.Element("movie")
    title = movie.get("title") or movie.get("original_title") or ""
    ET.SubElement(root, "title").text = title
    ET.SubElement(root, "originaltitle").text = movie.get("original_title", title)

    release = movie.get("release_date", "") or ""
    if release:
        ET.SubElement(root, "premiered").text = release
    ET.SubElement(root, "year").text = release[:4] if release else ""

    ET.SubElement(root, "plot").text = movie.get("overview", "") or ""

    rating = movie.get("vote_average")
    if rating is not None:
        ET.SubElement(root, "rating").text = f"{rating:.1f}"

    tmdb_id = movie.get("id")
    if tmdb_id:
        uid = ET.SubElement(root, "uniqueid", type="tmdb")
        uid.set("default", "true")
        uid.text = str(tmdb_id)

    cast = (movie.get("credits", {}).get("cast") or [])[:actors_count]
    for person in cast:
        actor_el = ET.SubElement(root, "actor")
        ET.SubElement(actor_el, "name").text = person.get("name", "") or ""
        ET.SubElement(actor_el, "role").text = person.get("character", "") or ""

    try:
        ET.indent(root, space="  ")
    except AttributeError:
        pass
    ET.ElementTree(root).write(nfo_path, encoding="utf-8", xml_declaration=True)
    return nfo_path


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class FetchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TMDB Poster & NFO Fetcher")
        self.geometry("1500x860")
        self.minsize(1200, 600)

        self.row_data: list[dict] = []
        self.worker: threading.Thread | None = None
        self.stop_flag = threading.Event()
        self.msg_queue: queue.Queue = queue.Queue()

        self.api_key_var    = tk.StringVar()
        self.language_var   = tk.StringVar(value="en-US")
        self.image_size_var = tk.StringVar(value="w500")
        self.actors_var     = tk.IntVar(value=10)
        self.adult_var      = tk.BooleanVar(value=True)
        self.progress_var   = tk.DoubleVar(value=0.0)
        self.status_var     = tk.StringVar(value="Ready.")

        self._build_ui()
        self._load_config()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_queue()

    # ==================================================================
    # UI construction
    # ==================================================================

    def _build_ui(self) -> None:
        self.rowconfigure(0, weight=0)
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=5)
        self.rowconfigure(3, weight=0)
        self.rowconfigure(4, weight=1)
        self.columnconfigure(0, weight=1)

        self._build_settings()
        self._build_toolbar()
        self._build_tree()
        self._build_action_bar()
        self._build_log()

    def _build_settings(self) -> None:
        f = ttk.Frame(self, padding=(10, 8, 10, 4))
        f.grid(row=0, column=0, sticky="ew")

        ttk.Label(f, text="TMDB API key:").pack(side=tk.LEFT)
        self.entry_api = ttk.Entry(f, textvariable=self.api_key_var, width=40)
        self.entry_api.pack(side=tk.LEFT, padx=(6, 16))

        ttk.Label(f, text="Language:").pack(side=tk.LEFT)
        ttk.Combobox(f, textvariable=self.language_var, values=LANGUAGES,
                     width=8, state="readonly").pack(side=tk.LEFT, padx=(6, 16))

        ttk.Label(f, text="Poster size:").pack(side=tk.LEFT)
        ttk.Combobox(f, textvariable=self.image_size_var, values=IMAGE_SIZES,
                     width=9, state="readonly").pack(side=tk.LEFT, padx=(6, 16))

        ttk.Label(f, text="Actors:").pack(side=tk.LEFT)
        ttk.Spinbox(f, from_=1, to=50, textvariable=self.actors_var,
                    width=4).pack(side=tk.LEFT, padx=(6, 16))

        ttk.Checkbutton(f, text="Include adult results",
                        variable=self.adult_var).pack(side=tk.LEFT)

    def _build_toolbar(self) -> None:
        f = ttk.Frame(self, padding=(10, 4, 10, 4))
        f.grid(row=1, column=0, sticky="ew")

        self.btn_add = ttk.Button(f, text="➕ Add Folder…",
                                  command=self._add_folder_dialog)
        self.btn_add.pack(side=tk.LEFT)

        self.btn_add_dir = ttk.Button(f, text="📁 Add All Subfolders…",
                                      command=self._add_directory_dialog)
        self.btn_add_dir.pack(side=tk.LEFT, padx=(6, 0))

        self.btn_remove = ttk.Button(f, text="➖ Remove Selected",
                                     command=self._remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=(6, 0))

        self.btn_clear = ttk.Button(f, text="🗑 Clear All",
                                    command=self._clear_all)
        self.btn_clear.pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(f, text="   ").pack(side=tk.LEFT)
        ttk.Label(f, textvariable=self.status_var,
                  foreground="gray").pack(side=tk.LEFT, padx=(20, 0))

    def _build_tree(self) -> None:
        f = ttk.Frame(self, padding=(10, 4, 10, 4))
        f.grid(row=2, column=0, sticky="nsew")
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        cols = ("idx", "folder", "query", "year", "match",
                "tmdb", "poster", "nfo", "status")
        self.tree = ttk.Treeview(f, columns=cols, show="headings",
                                 selectmode="extended")
        headings = {
            "idx":    ("#",              40,  "e"),
            "folder": ("Folder",        240,  "w"),
            "query":  ("Query (Title)", 220,  "w"),
            "year":   ("Year",           55, "center"),
            "match":  ("Matched (Title)",260, "w"),
            "tmdb":   ("TMDB ID",        85, "center"),
            "poster": ("Poster",         60, "center"),
            "nfo":    ("NFO",            60, "center"),
            "status": ("Status",        240,  "w"),
        }
        for c, (text, w, anchor) in headings.items():
            self.tree.heading(c, text=text)
            self.tree.column(c, width=w, anchor=anchor,
                             stretch=(c in ("folder", "match", "status")))

        self.tree.tag_configure("done",    background="#e6f4ea")
        self.tree.tag_configure("matched", background="#e8f0fe")
        self.tree.tag_configure("error",   background="#fdecea")
        self.tree.tag_configure("warn",    background="#fff4e5")

        vsb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(f, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Double-1>", self._on_double_click)

        self._ctx_menu = tk.Menu(self, tearoff=0)
        self._ctx_menu.add_command(label="✏️ Edit query…",
                                   command=self._on_edit_query)
        self._ctx_menu.add_command(label="✏️ Set TMDB ID…",
                                   command=self._on_set_tmdb_id)
        self._ctx_menu.add_command(label="🔄 Re-process this row",
                                   command=self._on_reprocess_row)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="📂 Open folder",
                                   command=self._on_open_folder)
        self._ctx_menu.add_command(label="➖ Remove from list",
                                   command=self._remove_selected)

    def _build_action_bar(self) -> None:
        f = ttk.Frame(self, padding=(10, 6, 10, 4))
        f.grid(row=3, column=0, sticky="ew")

        self.btn_search = ttk.Button(
            f, text="🔍 Search All",
            command=lambda: self._start_worker("search"))
        self.btn_search.pack(side=tk.LEFT)

        self.btn_download = ttk.Button(
            f, text="⬇️ Download All",
            command=lambda: self._start_worker("download"))
        self.btn_download.pack(side=tk.LEFT, padx=(6, 0))

        self.btn_process = ttk.Button(
            f, text="▶️ Process All (search + download)",
            command=lambda: self._start_worker("all"))
        self.btn_process.pack(side=tk.LEFT, padx=(6, 0))

        self.btn_stop = ttk.Button(
            f, text="⏹ Stop", command=self._stop_worker, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=(6, 0))

        self.progress = ttk.Progressbar(f, variable=self.progress_var,
                                        maximum=100, length=220)
        self.progress.pack(side=tk.RIGHT)

    def _build_log(self) -> None:
        f = ttk.LabelFrame(self, text="Log", padding=6)
        f.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        self.log_text = tk.Text(f, height=6, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(f, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")

    # ==================================================================
    # Config / shutdown
    # ==================================================================

    def _load_config(self) -> None:
        cfg = load_config()
        self.api_key_var.set(cfg.get("api_key", ""))
        self.language_var.set(cfg.get("language", "en-US"))
        self.image_size_var.set(cfg.get("image_size", "w500"))
        self.actors_var.set(int(cfg.get("actors", 10)))
        self.adult_var.set(bool(cfg.get("adult", True)))

    def _save_config(self) -> None:
        save_config({
            "api_key":    self.api_key_var.get().strip(),
            "language":   self.language_var.get(),
            "image_size": self.image_size_var.get(),
            "actors":     int(self.actors_var.get()),
            "adult":      bool(self.adult_var.get()),
        })

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("Quit",
                                       "A process is running. Stop it and quit?"):
                return
            self.stop_flag.set()
            self.worker.join(timeout=3)
        self._save_config()
        self.destroy()

    # ==================================================================
    # Folder management
    # ==================================================================

    def _add_folder_path(self, path: Path) -> None:
        if any(r["path"] == path for r in self.row_data):
            return
        cleaned = clean_and_extract(path.name)
        row = {
            "path":        path,
            "query_title": cleaned["title"],
            "query_year":  cleaned["year"] or "",
            "match":       "",
            "tmdb_id":     "",
            "override_id": "",
            "poster":      "",
            "nfo":         "",
            "status":      "Pending",
        }
        iid = self.tree.insert("", "end", values=(
            "", path.name,
            cleaned["title"], cleaned["year"] or "",
            "", "", "", "", "Pending",
        ))
        row["iid"] = iid
        self.row_data.append(row)
        self._renumber()

    def _renumber(self) -> None:
        for i, row in enumerate(self.row_data):
            vals = list(self.tree.item(row["iid"], "values"))
            vals[0] = str(i + 1)
            self.tree.item(row["iid"], values=vals)

    def _add_folder_dialog(self) -> None:
        d = filedialog.askdirectory(title="Select a movie folder")
        if d:
            self._add_folder_path(Path(d))

    def _add_directory_dialog(self) -> None:
        d = filedialog.askdirectory(
            title="Select a parent folder — its immediate subfolders "
                  "will be added")
        if not d:
            return
        parent = Path(d)
        subs = sorted(p for p in parent.iterdir() if p.is_dir())
        if not subs:
            messagebox.showinfo("No subfolders",
                                f"No subfolders found inside:\n{parent}")
            return
        if not messagebox.askyesno(
            "Add subfolders",
            f"Add {len(subs)} subfolders of:\n{parent}\n\nContinue?"):
            return
        for p in subs:
            self._add_folder_path(p)
        self._log(f"Added {len(subs)} subfolders from {parent}")

    def _remove_selected(self) -> None:
        for iid in self.tree.selection():
            for i, row in enumerate(self.row_data):
                if row["iid"] == iid:
                    self.tree.delete(iid)
                    self.row_data.pop(i)
                    break
        self._renumber()

    def _clear_all(self) -> None:
        if not self.row_data:
            return
        if not messagebox.askyesno("Clear all",
                                   "Remove all rows from the list?"):
            return
        for row in self.row_data:
            self.tree.delete(row["iid"])
        self.row_data.clear()
        self.progress_var.set(0)
        self.status_var.set("Ready.")

    # ==================================================================
    # Worker thread
    # ==================================================================

    def _start_worker(self, mode: str) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Busy", "A process is already running.")
            return
        if not self.row_data:
            messagebox.showinfo("No folders",
                                "Please add at least one folder first.")
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showerror(
                "API key required",
                "Please enter your TMDB API key at the top of the window.")
            return

        settings = {
            "api_key":       api_key,
            "language":      self.language_var.get(),
            "image_size":    self.image_size_var.get(),
            "actors":        int(self.actors_var.get()),
            "include_adult": bool(self.adult_var.get()),
        }

        self.stop_flag.clear()
        self.progress_var.set(0)
        self._set_running(True)
        self._log(f"Starting '{mode}' for {len(self.row_data)} folders…")

        self.worker = threading.Thread(
            target=self._worker, args=(mode, settings), daemon=True)
        self.worker.start()

    def _stop_worker(self) -> None:
        self.stop_flag.set()
        self._log("Stop requested — finishing current item…")

    def _worker(self, mode: str, settings: dict) -> None:
        try:
            if mode in ("search", "all"):
                self._search_phase(settings)
            if mode in ("download", "all") and not self.stop_flag.is_set():
                self._download_phase(settings)
        except Exception as exc:  # noqa: BLE001
            self.msg_queue.put({"type": "log",
                                "text": f"Worker error: {exc!r}"})
        finally:
            self.msg_queue.put({"type": "done", "mode": mode})

    def _search_phase(self, settings: dict) -> None:
        total = len(self.row_data)
        for i, row in enumerate(self.row_data):
            if self.stop_flag.is_set():
                return
            self.msg_queue.put({"type": "progress", "current": i + 1,
                                "total": total})
            if row.get("override_id") or row.get("tmdb_id"):
                continue

            title = row.get("query_title") or ""
            year  = row.get("query_year") or None

            self.msg_queue.put({"type": "row", "index": i,
                                "fields": {"status": "Searching…"}})
            self.msg_queue.put({"type": "log",
                                "text": f"→ {row['path'].name}"})

            hit, used_v, used_y = search_movie_robust(
                title, year,
                settings["api_key"], settings["language"],
                settings["include_adult"],
                log_cb=lambda t: self.msg_queue.put(
                    {"type": "log", "text": t}),
            )

            if hit:
                year_r = (hit.get("release_date") or "")[:4]
                title_r = hit.get("title") or hit.get("original_title") or "?"
                match_str = f"{title_r} ({year_r})" if year_r else title_r
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {
                                        "tmdb_id": str(hit["id"]),
                                        "match":   match_str,
                                        "status":  "Matched",
                                    }})
                self.msg_queue.put({"type": "log",
                                    "text": f"✓ {row['path'].name} "
                                            f"→ {match_str}"})
            else:
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"status": "No match"}})
                self.msg_queue.put({"type": "log",
                                    "text": f"? No match: {row['path'].name} "
                                            f"(title={title!r} year={year!r})"})
            time.sleep(0.15)

    def _download_phase(self, settings: dict) -> None:
        total = len(self.row_data)
        for i, row in enumerate(self.row_data):
            if self.stop_flag.is_set():
                return
            self.msg_queue.put({"type": "progress", "current": i + 1,
                                "total": total})

            tmdb_id = row.get("override_id") or row.get("tmdb_id")
            if not tmdb_id:
                continue

            try:
                details = tmdb_details(int(tmdb_id), settings["api_key"],
                                       settings["language"])
            except Exception as exc:  # noqa: BLE001
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"status": f"Details error: {exc}"}})
                self.msg_queue.put({"type": "log",
                                    "text": f"✗ {row['path'].name}: {exc}"})
                time.sleep(0.25)
                continue

            pp = details.get("poster_path")
            if pp:
                try:
                    download_poster(pp, row["path"] / "poster.jpg",
                                    settings["image_size"])
                    self.msg_queue.put({"type": "row", "index": i,
                                        "fields": {"poster": "✅"}})
                except Exception as exc:  # noqa: BLE001
                    self.msg_queue.put({"type": "row", "index": i,
                                        "fields": {"poster": "❌"}})
                    self.msg_queue.put({"type": "log",
                                        "text": f"✗ Poster: {exc}"})
            else:
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"poster": "—"}})

            try:
                nfo = write_nfo(details, row["path"], settings["actors"])
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"nfo": "✅",
                                               "status": f"Done → {nfo.name}"}})
                self.msg_queue.put({"type": "log",
                                    "text": f"✓ {row['path'].name} saved "
                                            f"{nfo.name}"})
            except Exception as exc:  # noqa: BLE001
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"nfo": "❌",
                                               "status": f"NFO error: {exc}"}})
                self.msg_queue.put({"type": "log",
                                    "text": f"✗ NFO: {exc}"})
            time.sleep(0.25)

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        for b in (self.btn_search, self.btn_download, self.btn_process,
                  self.btn_add, self.btn_add_dir, self.btn_remove,
                  self.btn_clear):
            b.config(state=state)
        self.btn_stop.config(state="normal" if running else "disabled")
        for w in (self.entry_api,):
            w.config(state="disabled" if running else "normal")
        if not running:
            self.status_var.set("Ready.")

    # ==================================================================
    # Queue polling & row updates
    # ==================================================================

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                self._handle_message(msg)
        except queue.Empty:
            pass
        self.after(80, self._poll_queue)

    def _handle_message(self, msg: dict) -> None:
        t = msg["type"]
        if t == "log":
            self._log(msg["text"])
        elif t == "row":
            self._apply_row_update(msg["index"], msg["fields"])
        elif t == "progress":
            cur, tot = msg["current"], msg["total"]
            pct = (cur / tot * 100) if tot else 0
            self.progress_var.set(pct)
            self.status_var.set(f"{cur}/{tot} ({pct:.0f}%)")
        elif t == "done":
            self._set_running(False)
            self._log(f"Finished '{msg['mode']}'.")
            if not self.stop_flag.is_set():
                self.progress_var.set(100)

    def _apply_row_update(self, i: int, fields: dict) -> None:
        if not (0 <= i < len(self.row_data)):
            return
        row = self.row_data[i]
        row.update(fields)

        status = row.get("status", "")
        if status.startswith("Done"):
            tag = "done"
        elif status.startswith("Matched"):
            tag = "matched"
        elif ("error" in status.lower()) or status.startswith("❌"):
            tag = "error"
        elif status in ("No match",) or status.startswith("Query edited"):
            tag = "warn"
        else:
            tag = ""

        self.tree.item(
            row["iid"], tags=(tag,) if tag else (),
            values=(
                i + 1,
                row["path"].name,
                row.get("query_title", ""),
                row.get("query_year", ""),
                row.get("match", ""),
                row.get("override_id") or row.get("tmdb_id", ""),
                row.get("poster", ""),
                row.get("nfo", ""),
                status,
            ),
        )

    # ==================================================================
    # Context menu actions
    # ==================================================================

    def _on_right_click(self, event) -> None:
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _on_double_click(self, event) -> None:
        col = self.tree.identify_column(event.x)
        if col == "#3":          # Query column
            self._on_edit_query()
        elif col == "#4":        # Year column
            self._on_edit_query()
        elif col == "#6":        # TMDB ID column
            self._on_set_tmdb_id()

    def _selected_row(self) -> tuple[int, dict] | tuple[None, None]:
        sel = self.tree.selection()
        if not sel:
            return None, None
        iid = sel[0]
        for i, r in enumerate(self.row_data):
            if r["iid"] == iid:
                return i, r
        return None, None

    def _on_edit_query(self) -> None:
        i, row = self._selected_row()
        if row is None:
            return

        # Dialog with two fields
        dlg = tk.Toplevel(self)
        dlg.title("Edit query")
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Folder:").grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text=row["path"].name,
                  wraplength=380).grid(row=0, column=1, sticky="w",
                                       columnspan=2, pady=(0, 8))

        ttk.Label(frm, text="Search title:").grid(row=1, column=0, sticky="w")
        title_var = tk.StringVar(value=row.get("query_title", ""))
        ttk.Entry(frm, textvariable=title_var, width=42).grid(
            row=1, column=1, sticky="ew", columnspan=2, pady=(0, 6))

        ttk.Label(frm, text="Year (optional):").grid(row=2, column=0, sticky="w")
        year_var = tk.StringVar(value=row.get("query_year", ""))
        ttk.Entry(frm, textvariable=year_var, width=10).grid(
            row=2, column=1, sticky="w", pady=(0, 6))

        ttk.Label(
            frm,
            text=("Variants that will be tried:\n"
                  "• title with hyphen → space\n"
                  "• title with hyphens/apostrophes removed\n"
                  "• title with colons removed\n"
                  "Each tried with year first, then without."),
            foreground="gray",
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 6))

        auto_search_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Search this row after saving",
                        variable=auto_search_var).grid(
            row=4, column=0, columnspan=3, sticky="w")

        def _ok():
            new_title = title_var.get().strip()
            new_year  = year_var.get().strip()
            if new_year and not new_year.isdigit():
                messagebox.showerror("Invalid year",
                                     "Year must be numeric (e.g. 2015).",
                                     parent=dlg)
                return
            row["query_title"] = new_title
            row["query_year"]  = new_year
            # Clear prior match so a re-search will actually run
            row["match"] = ""
            row["tmdb_id"] = ""
            row["poster"] = ""
            row["nfo"] = ""
            row["status"] = "Query edited — re-search to apply"
            self._apply_row_update(i, {})
            self._log(f"Query for {row['path'].name}: "
                      f"title={new_title!r} year={new_year!r}")
            dlg.destroy()

            if auto_search_var.get() and new_title:
                self._reprocess_row_async(i)

        def _cancel():
            dlg.destroy()

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Cancel", command=_cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Save", command=_ok).pack(side=tk.RIGHT, padx=(0, 6))

        dlg.wait_window()

    def _on_set_tmdb_id(self) -> None:
        i, row = self._selected_row()
        if row is None:
            return
        new_id = simpledialog.askstring(
            "Override TMDB ID",
            f"Set TMDB ID for:\n{row['path'].name}\n\n"
            f"(Leave blank to clear the override.)",
            initialvalue=row.get("override_id", ""), parent=self)
        if new_id is None:
            return
        new_id = new_id.strip()
        if new_id and not new_id.isdigit():
            messagebox.showerror("Invalid ID", "TMDB ID must be a number.")
            return
        row["override_id"] = new_id
        row["status"] = ("Override set — re-process to apply"
                         if new_id else "Override cleared")
        self._apply_row_update(i, {})
        self._log(f"Override for {row['path'].name}: {new_id or '(cleared)'}")

    def _on_reprocess_row(self) -> None:
        i, row = self._selected_row()
        if row is None:
            return
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Busy",
                                   "Wait for the current run to finish.")
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("API key required",
                                 "Please enter your TMDB API key.")
            return
        self._reprocess_row_async(i)

    def _reprocess_row_async(self, i: int) -> None:
        settings = {
            "api_key":       self.api_key_var.get().strip(),
            "language":      self.language_var.get(),
            "image_size":    self.image_size_var.get(),
            "actors":        int(self.actors_var.get()),
            "include_adult": bool(self.adult_var.get()),
        }
        self.stop_flag.clear()
        self._set_running(True)
        self.worker = threading.Thread(
            target=self._reprocess_single, args=(i, settings), daemon=True)
        self.worker.start()

    def _reprocess_single(self, i: int, settings: dict) -> None:
        try:
            row = self.row_data[i]
            tmdb_id = row.get("override_id") or row.get("tmdb_id")

            if not tmdb_id:
                title = row.get("query_title") or ""
                year  = row.get("query_year") or None
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"status": "Searching…"}})
                self.msg_queue.put({"type": "log",
                                    "text": f"→ {row['path'].name}"})
                hit, _, _ = search_movie_robust(
                    title, year,
                    settings["api_key"], settings["language"],
                    settings["include_adult"],
                    log_cb=lambda t: self.msg_queue.put(
                        {"type": "log", "text": t}))
                if not hit:
                    self.msg_queue.put({"type": "row", "index": i,
                                        "fields": {"status": "No match"}})
                    return
                tmdb_id = hit["id"]
                year_r = (hit.get("release_date") or "")[:4]
                title_r = hit.get("title") or hit.get("original_title") or "?"
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {
                                        "tmdb_id": str(tmdb_id),
                                        "match": f"{title_r} ({year_r})"
                                                 if year_r else title_r,
                                        "status": "Matched",
                                    }})
                time.sleep(0.15)

            details = tmdb_details(int(tmdb_id), settings["api_key"],
                                   settings["language"])
            pp = details.get("poster_path")
            if pp:
                try:
                    download_poster(pp, row["path"] / "poster.jpg",
                                    settings["image_size"])
                    self.msg_queue.put({"type": "row", "index": i,
                                        "fields": {"poster": "✅"}})
                except Exception as exc:  # noqa: BLE001
                    self.msg_queue.put({"type": "row", "index": i,
                                        "fields": {"poster": "❌"}})
                    self.msg_queue.put({"type": "log",
                                        "text": f"✗ Poster: {exc}"})
            try:
                nfo = write_nfo(details, row["path"], settings["actors"])
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"nfo": "✅",
                                               "status": f"Done → {nfo.name}"}})
            except Exception as exc:  # noqa: BLE001
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"nfo": "❌",
                                               "status": f"NFO error: {exc}"}})
        finally:
            self.msg_queue.put({"type": "done", "mode": "reprocess"})

    def _on_open_folder(self) -> None:
        _, row = self._selected_row()
        if row is None:
            return
        folder = row["path"]
        try:
            if sys.platform == "win32":
                os.startfile(str(folder))  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Cannot open folder", str(exc))

    # ==================================================================
    # Log
    # ==================================================================

    def _log(self, text: str) -> None:
        ts = time.strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{ts}] {text}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def main() -> None:
    FetchApp().mainloop()


if __name__ == "__main__":
    main()