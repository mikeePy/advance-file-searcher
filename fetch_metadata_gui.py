"""
fetch_metadata_gui.py  (skip-if-complete edition)
==================================================
Desktop GUI for fetching posters and NFO metadata from TMDB.
Handles BOTH movies and TV series via a batch-level type toggle.

New in this version
-------------------
• "Skip folders that already have poster.jpg + .nfo" checkbox
  (default ON). When adding folders — single or bulk — any
  folder that already contains BOTH a poster.jpg AND at least
  one .nfo file is skipped and NOT added to the process list.
• Add-skipped folders are counted and logged.
• Trying to single-add a folder that's already complete shows
  an information dialog so you know why nothing happened.
• The check is case-insensitive (Poster.JPG, MOVIE.NFO, etc.
  all count) so it works on both Windows and Linux.

Still included
--------------
• Season / S0x / COMPLETE / 第X季 / 全XX集 markers stripped
• Year verification (prefers results matching the folder's year)
• Batch media-type toggle: Movie or TV
• Manual query editor and TMDB ID override (right-click)
• Background worker thread, progress bar, live log

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
# Config
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

TMDB_SEARCH_MOVIE = "https://api.themoviedb.org/3/search/movie"
TMDB_SEARCH_TV    = "https://api.themoviedb.org/3/search/tv"
TMDB_MOVIE_URL    = "https://api.themoviedb.org/3/movie/{id}"
TMDB_TV_URL       = "https://api.themoviedb.org/3/tv/{id}"
TMDB_IMAGE_URL    = "https://image.tmdb.org/t/p/{size}{path}"

TV_HINT_RE = re.compile(
    r"\b(?:"
    r"season\s*\d+|s\d{1,2}(?:e\d{1,2})?|"
    r"complete|全\s*\d*\s*集|全\s*ep|episode\s*\d+|"
    r"ep\s*\d+|第\s*\d+\s*[季集話话]|"
    r"tv\s*series|tv\s*show|drama\s*series"
    r")\b",
    re.I,
)

RELEASE_GROUPS = [
    r"YTS(?:\.MX|\.AG|\.AM|\.PE|\.LT)?", r"YIFY",
    r"VXT(?:-AcEs)?", r"AcEs", r"Ganool(?:\.?iconvid\.co)?",
    r"CMCT", r"BYNDR", r"Flux", r"GalaxyRG",
    r"Tv21\.org", r"iconvid\.co",
    r"Vegamovies(?:\.is|\.hot|\.com)?",
    r"scOrp[_\s]*sujaidr", r"BOKUTOX", r"Deceit", r"ETRG",
    r"P2P", r"Fugi", r"ZhuixinFan(?:V2)?",
    r"ENG[\.\s]*HC", r"ENGHC", r"AEF",
    r"x264-[A-Za-z0-9]+", r"x265-[A-Za-z0-9]+",
    r"Crazy4ad", r"ShAaNiG", r"MkvCage", r"GalaxyTV",
    r"\[TGx\]", r"\[ettv\]", r"\[eztv\.re\]", r"\[i_c\]",
    r"\[y2flix\]", r"\[GWC\]", r"\[fs87\]", r"\[PSA\]",
    r"\[TorrentCouch\.com\]", r"\[HDzone\]", r"\[NOP\]",
    r"ION10", r"V3SP4EV3R", r"EliteT", r"STRiFE",
    r"MiNX", r"PHOENiX", r"MEMENTO", r"TBS",
]

NOISE_TOKENS = (
    r"720p|1080p|2160p|480p|360p|4k|uhd|hd|"
    r"bluray|blu-ray|brrip|bdrip|webrip|web-dl|webdl|web|"
    r"hdrip|dvdrip|hdtv|dvd|amzn|dsnp|nf|hulu|"
    r"x264|x265|h\.?264|h\.?265|hevc|"
    r"aac|ac3|dts|ddp\d?|dd\d|atmos|"
    r"10bit|8bit|5\.1|7\.1|2\.0|2ch|"
    r"repack|proper|extended|remastered|directors?\.?cut|"
    r"hc|eng|engsub|engsubs|subs|subbed|subtitles|esubs|e-subs|"
    r"multi|multisub|multi-sub|japanese|chinese|jap|"
    r"hin-jap|hindi|english|dual\.?audio|dubbed|"
    r"cineby|kisskh|mkv|mp4|avi|"
    r"vol\.?\s*\d+|disc\s*\d+|d\d+|"
    r"ita|jpn|kor|chi|sub"
)


# ---------------------------------------------------------------------------
# Query cleaner
# ---------------------------------------------------------------------------


def extract_years(text: str) -> list[str]:
    return re.findall(r"\b(19\d{2}|20\d{2})\b", text)


def looks_like_tv(name: str) -> bool:
    return bool(TV_HINT_RE.search(name))


def clean_and_extract(name: str) -> dict:
    """Turn a messy folder name into {title, year, variants}."""
    s = name

    s = re.sub(r"\[[^\]]*\]", " ", s)
    s = re.sub(r"【[^】]*】", " ", s)
    s = re.sub(r"\{[^}]*\}", " ", s)

    s = re.sub(
        r"\([^)]*(?:YTS|BluRay|WEB|HDRip|BRRip|DVDrip|720p|1080p|480p|"
        r"2160p|4K|UHD|AAC|x264|x265|H264|H265|DDP|HDTV|Remastered|"
        r"Extended|Directors?\.?\s*Cut)[^)]*\)",
        " ", s, flags=re.I,
    )

    s = re.sub(r"^(?:720p|1080p|2160p|4k)\s+(?:hd\s+)?movie\s+",
               "", s, flags=re.I)

    s = s.replace(".", " ").replace("_", " ")

    for grp in RELEASE_GROUPS:
        s = re.sub(grp, " ", s, flags=re.I)

    s = re.sub(r"-\s*[A-Za-z0-9_]{2,15}\s*$", " ", s)

    s = re.sub(r"\b(?:" + NOISE_TOKENS + r")\b", " ", s, flags=re.I)

    # Season / complete markers
    s = re.sub(r"\bseason\s*\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\bcomplete\b", " ", s, flags=re.I)
    s = re.sub(r"全\s*EP\s*\d+", " ", s, flags=re.I)
    s = re.sub(r"全\s*\d+\s*集", " ", s)
    s = re.sub(r"\b\d+\s*集全\b", " ", s)
    s = re.sub(r"第\s*\d+\s*[季集話话]", " ", s)
    s = re.sub(r"\b\d{1,3}\s*-\s*\d{1,3}\s*完\b", " ", s)

    # Episode markers
    s = re.sub(r"\bepisode\s+\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\bep\.?\s*\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\bs\d{1,2}\s*e\d{1,2}\b", " ", s, flags=re.I)
    s = re.sub(r"\bs\d{1,2}\b", " ", s, flags=re.I)

    s = re.sub(r"\s+", " ", s).strip(" -.:;,·|/")

    years = extract_years(s)
    year = years[-1] if years else None

    if year:
        idx = s.rfind(year)
        title = (s[:idx] + s[idx + len(year):]).strip(" -.:;,·|/")
        if not title:
            title = year
    else:
        title = s

    variants: list[str] = []
    if title:
        variants.append(title)
        variants.append(re.sub(r"\s+", " ", title.replace("-", " ")).strip())
        variants.append(title.replace("-", ""))
        variants.append(title.replace("'", "").replace("’", ""))
        variants.append(title.replace(":", ""))
        variants.append(re.sub(r"\s+", " ", title.replace("-", " ")
                                .replace("'", "").replace("’", "")).strip())

    seen, unique = set(), []
    for v in variants:
        v = v.strip()
        if v and v not in seen:
            seen.add(v)
            unique.append(v)

    return {"title": title, "year": year, "variants": unique}


def build_variants(title: str) -> list[str]:
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
# Folder completeness check
# ---------------------------------------------------------------------------


def has_complete_metadata(folder: Path) -> bool:
    """
    True if the folder already contains BOTH:
      • a file called poster.jpg (case-insensitive), AND
      • at least one .nfo file anywhere in that folder.
    """
    has_poster = False
    has_nfo = False
    try:
        for p in folder.iterdir():
            if not p.is_file():
                continue
            low = p.name.lower()
            if low == "poster.jpg":
                has_poster = True
            elif low.endswith(".nfo"):
                has_nfo = True
            if has_poster and has_nfo:
                return True
    except (PermissionError, FileNotFoundError):
        return False
    return has_poster and has_nfo


# ---------------------------------------------------------------------------
# TMDB calls
# ---------------------------------------------------------------------------


def search_robust(
    title: str,
    year: str | None,
    media_type: str,
    api_key: str,
    language: str,
    include_adult: bool,
    log_cb=None,
):
    variants = build_variants(title)
    if not variants:
        return None, None, None

    if media_type == "tv":
        endpoint = TMDB_SEARCH_TV
        year_param = "first_air_date_year"
        date_field = "first_air_date"
    else:
        endpoint = TMDB_SEARCH_MOVIE
        year_param = "year"
        date_field = "release_date"

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
            params[year_param] = y
        try:
            r = requests.get(endpoint, params=params, timeout=15)
            if r.status_code == 429:
                if log_cb:
                    log_cb("  ! rate-limited, sleeping 2 s")
                time.sleep(2)
                continue
            r.raise_for_status()
            results = r.json().get("results", [])
        except Exception as exc:
            if log_cb:
                log_cb(f"  ! query={v!r}: {exc}")
            time.sleep(0.1)
            continue

        if not results:
            if log_cb:
                log_cb(f"  · [{media_type}] query={v!r} "
                       f"{year_param}={y!r} → 0 results")
            time.sleep(0.1)
            continue

        if y:
            matching = [
                r for r in results
                if (r.get(date_field) or "").startswith(y)
            ]
            if matching:
                if log_cb:
                    name = (matching[0].get("title")
                            or matching[0].get("name"))
                    log_cb(f"  ✓ [{media_type}] query={v!r} "
                           f"{year_param}={y!r} → {name} (year match)")
                return matching[0], v, y
            else:
                top_year = (results[0].get(date_field) or "")[:4]
                if log_cb:
                    name = (results[0].get("title")
                            or results[0].get("name"))
                    log_cb(f"  ⚠ [{media_type}] query={v!r} "
                           f"{year_param}={y!r} → {name} "
                           f"(year mismatch: found {top_year or '?'})")
                return results[0], v, y
        else:
            if log_cb:
                name = results[0].get("title") or results[0].get("name")
                log_cb(f"  ✓ [{media_type}] query={v!r} → {name}")
            return results[0], v, None

    return None, None, None


def tmdb_details(movie_id: int, media_type: str,
                 api_key: str, language: str) -> dict:
    if media_type == "tv":
        url = TMDB_TV_URL.format(id=movie_id)
        append = "credits,external_ids"
    else:
        url = TMDB_MOVIE_URL.format(id=movie_id)
        append = "credits"
    params = {
        "api_key": api_key,
        "language": language,
        "append_to_response": append,
    }
    r = requests.get(url, params=params, timeout=15)
    r.raise_for_status()
    return r.json()


def download_poster(poster_path: str, dest: Path, image_size: str) -> None:
    url = TMDB_IMAGE_URL.format(size=image_size, path=poster_path)
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    if len(r.content) < 1000:
        raise ValueError(f"Poster too small ({len(r.content)} bytes)")
    dest.write_bytes(r.content)


# ---------------------------------------------------------------------------
# NFO writers
# ---------------------------------------------------------------------------


def write_movie_nfo(movie: dict, folder: Path, actors_count: int) -> Path:
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


def write_tv_nfo(show: dict, folder: Path, actors_count: int) -> Path:
    nfo_path = folder / f"{folder.name}.nfo"

    root = ET.Element("tvshow")
    name = show.get("name") or show.get("original_name") or ""
    ET.SubElement(root, "title").text = name
    ET.SubElement(root, "originaltitle").text = show.get("original_name", name)

    first_air = show.get("first_air_date", "") or ""
    if first_air:
        ET.SubElement(root, "premiered").text = first_air
    ET.SubElement(root, "year").text = first_air[:4] if first_air else ""
    ET.SubElement(root, "plot").text = show.get("overview", "") or ""

    rating = show.get("vote_average")
    if rating is not None:
        ET.SubElement(root, "rating").text = f"{rating:.1f}"

    tmdb_id = show.get("id")
    if tmdb_id:
        uid = ET.SubElement(root, "uniqueid", type="tmdb")
        uid.set("default", "true")
        uid.text = str(tmdb_id)

    cast = (show.get("credits", {}).get("cast") or [])[:actors_count]
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


def write_nfo(details: dict, folder: Path, media_type: str,
              actors_count: int) -> Path:
    if media_type == "tv":
        return write_tv_nfo(details, folder, actors_count)
    return write_movie_nfo(details, folder, actors_count)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class FetchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TMDB Poster & NFO Fetcher  —  Batch Movies / TV")
        self.geometry("1550x900")
        self.minsize(1200, 620)

        self.row_data: list[dict] = []
        self.worker: threading.Thread | None = None
        self.stop_flag = threading.Event()
        self.msg_queue: queue.Queue = queue.Queue()

        self.api_key_var        = tk.StringVar()
        self.language_var       = tk.StringVar(value="en-US")
        self.image_size_var     = tk.StringVar(value="w500")
        self.actors_var         = tk.IntVar(value=10)
        self.adult_var          = tk.BooleanVar(value=True)
        self.batch_type_var     = tk.StringVar(value="movie")
        self.skip_existing_var  = tk.BooleanVar(value=True)
        self.progress_var       = tk.DoubleVar(value=0.0)
        self.status_var         = tk.StringVar(value="Ready.")

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
        self.rowconfigure(2, weight=0)
        self.rowconfigure(3, weight=5)
        self.rowconfigure(4, weight=0)
        self.rowconfigure(5, weight=1)
        self.columnconfigure(0, weight=1)

        self._build_api_bar()
        self._build_batch_bar()
        self._build_toolbar()
        self._build_tree()
        self._build_action_bar()
        self._build_log()

    def _build_api_bar(self) -> None:
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

    def _build_batch_bar(self) -> None:
        f = ttk.LabelFrame(self, text="Batch media type", padding=(10, 4, 10, 6))
        f.grid(row=1, column=0, sticky="ew", padx=10, pady=(2, 4))

        ttk.Label(
            f,
            text=("Choose what EVERY folder in this batch is. "
                  "Reality / scripted / anime / variety series all count as "
                  "TV Series on TMDB (e.g. Singles Inferno, 19th Floor, "
                  "Game of Thrones). Season / S0x / COMPLETE / 第X季 markers "
                  "are automatically stripped from the search query."),
            foreground="gray",
            wraplength=1200,
            justify="left",
        ).pack(anchor="w", pady=(0, 4))

        row = ttk.Frame(f)
        row.pack(fill=tk.X)

        ttk.Radiobutton(row, text="🎬 Movies",
                        variable=self.batch_type_var, value="movie",
                        command=self._on_batch_type_changed).pack(side=tk.LEFT)
        ttk.Radiobutton(row, text="📺 TV Series",
                        variable=self.batch_type_var, value="tv",
                        command=self._on_batch_type_changed).pack(
            side=tk.LEFT, padx=(12, 0))

        ttk.Button(row, text="↻ Apply this type to all loaded rows",
                   command=self._apply_batch_type_to_all).pack(
            side=tk.LEFT, padx=(24, 0))

        self.batch_hint = ttk.Label(row, text="", foreground="gray")
        self.batch_hint.pack(side=tk.LEFT, padx=(24, 0))

    def _build_toolbar(self) -> None:
        f = ttk.Frame(self, padding=(10, 4, 10, 4))
        f.grid(row=2, column=0, sticky="ew")

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

        ttk.Label(f, text="  |  ").pack(side=tk.LEFT)
        ttk.Checkbutton(
            f,
            text="Skip folders that already have poster.jpg + .nfo",
            variable=self.skip_existing_var,
        ).pack(side=tk.LEFT)

        ttk.Label(f, text="   ").pack(side=tk.LEFT)
        ttk.Label(f, textvariable=self.status_var,
                  foreground="gray").pack(side=tk.LEFT, padx=(20, 0))

    def _build_tree(self) -> None:
        f = ttk.Frame(self, padding=(10, 4, 10, 4))
        f.grid(row=3, column=0, sticky="nsew")
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)

        cols = ("idx", "folder", "type", "hint", "query", "year",
                "match", "tmdb", "poster", "nfo", "status")
        self.tree = ttk.Treeview(f, columns=cols, show="headings",
                                 selectmode="extended")
        headings = {
            "idx":    ("#",              40,  "e"),
            "folder": ("Folder",        230,  "w"),
            "type":   ("Type",           55, "center"),
            "hint":   ("Hint",           90, "center"),
            "query":  ("Query (Title)", 200,  "w"),
            "year":   ("Year",           50, "center"),
            "match":  ("Matched (Title)",220, "w"),
            "tmdb":   ("TMDB ID",        80, "center"),
            "poster": ("Poster",         55, "center"),
            "nfo":    ("NFO",            55, "center"),
            "status": ("Status",        200,  "w"),
        }
        for c, (text, w, anchor) in headings.items():
            self.tree.heading(c, text=text)
            self.tree.column(c, width=w, anchor=anchor,
                             stretch=(c in ("folder", "match", "status")))

        self.tree.tag_configure("done",    background="#e6f4ea")
        self.tree.tag_configure("matched", background="#e8f0fe")
        self.tree.tag_configure("error",   background="#fdecea")
        self.tree.tag_configure("warn",    background="#fff4e5")
        self.tree.tag_configure("hint",    foreground="#c47a00")

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
        self._ctx_menu.add_command(label="🔀 Toggle this row Movie / TV",
                                   command=self._on_toggle_media_type)
        self._ctx_menu.add_command(label="🔄 Re-process this row",
                                   command=self._on_reprocess_row)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="📂 Open folder",
                                   command=self._on_open_folder)
        self._ctx_menu.add_command(label="➖ Remove from list",
                                   command=self._remove_selected)

    def _build_action_bar(self) -> None:
        f = ttk.Frame(self, padding=(10, 6, 10, 4))
        f.grid(row=4, column=0, sticky="ew")

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
        f.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0, 10))
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
        self.batch_type_var.set(cfg.get("batch_type", "movie"))
        self.skip_existing_var.set(bool(cfg.get("skip_existing", True)))
        self._on_batch_type_changed()

    def _save_config(self) -> None:
        save_config({
            "api_key":       self.api_key_var.get().strip(),
            "language":      self.language_var.get(),
            "image_size":    self.image_size_var.get(),
            "actors":        int(self.actors_var.get()),
            "adult":         bool(self.adult_var.get()),
            "batch_type":    self.batch_type_var.get(),
            "skip_existing": bool(self.skip_existing_var.get()),
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
    # Batch-type bar
    # ==================================================================

    def _on_batch_type_changed(self) -> None:
        t = self.batch_type_var.get()
        label = "Movies" if t == "movie" else "TV Series"
        self.batch_hint.config(
            text=f"New folders will be added as: {label}"
        )

    def _apply_batch_type_to_all(self) -> None:
        if not self.row_data:
            messagebox.showinfo("No rows",
                                "Load some folders first.")
            return
        t = self.batch_type_var.get()
        label = "Movies" if t == "movie" else "TV Series"
        if not messagebox.askyesno(
            "Apply to all rows",
            f"Set ALL {len(self.row_data)} loaded rows to {label}?\n\n"
            "Any previous matches will be cleared so you can re-search."
        ):
            return
        for i, row in enumerate(self.row_data):
            row["media_type"] = t
            row["match"] = ""
            row["tmdb_id"] = ""
            row["poster"] = ""
            row["nfo"] = ""
            row["status"] = f"Type set to {label} — re-search to apply"
            self._apply_row_update(i, {})
        self._log(f"Applied batch type {label} to all {len(self.row_data)} rows.")

    # ==================================================================
    # Folder management
    # ==================================================================

    def _add_folder_path(self, path: Path, silent: bool = False) -> bool:
        """
        Add a folder to the list.
        Returns True if added, False if skipped (already in list,
        not a directory, or already has poster.jpg + .nfo).
        """
        if any(r["path"] == path for r in self.row_data):
            return False
        if not path.is_dir():
            return False

        if self.skip_existing_var.get() and has_complete_metadata(path):
            if not silent:
                self._log(f"⏭️ Skipped (already has poster.jpg + .nfo): "
                          f"{path.name}")
                messagebox.showinfo(
                    "Skipped",
                    f"This folder already has poster.jpg and .nfo:\n\n"
                    f"{path}\n\n"
                    f"It was NOT added to the process list.\n\n"
                    f"To force-add it, untick "
                    f"'Skip folders that already have poster.jpg + .nfo'.",
                )
            return False

        cleaned = clean_and_extract(path.name)
        media_type = self.batch_type_var.get()

        hint = ""
        if media_type == "movie" and looks_like_tv(path.name):
            hint = "looks like TV"
        elif media_type == "tv" and not looks_like_tv(path.name):
            hint = "check type"

        row = {
            "path":        path,
            "media_type":  media_type,
            "query_title": cleaned["title"],
            "query_year":  cleaned["year"] or "",
            "match":       "",
            "tmdb_id":     "",
            "override_id": "",
            "poster":      "",
            "nfo":         "",
            "status":      "Pending",
            "hint":        hint,
        }
        iid = self.tree.insert("", "end", values=(
            "", path.name, media_type.upper(), hint,
            cleaned["title"], cleaned["year"] or "",
            "", "", "", "", "Pending",
        ))
        row["iid"] = iid
        self.row_data.append(row)
        self._renumber()
        return True

    def _renumber(self) -> None:
        for i, row in enumerate(self.row_data):
            vals = list(self.tree.item(row["iid"], "values"))
            vals[0] = str(i + 1)
            self.tree.item(row["iid"], values=vals)

    def _add_folder_dialog(self) -> None:
        d = filedialog.askdirectory(title="Select a movie or TV folder")
        if not d:
            return
        path = Path(d)
        added = self._add_folder_path(path, silent=False)
        if added:
            self._log(f"Added: {path.name}")

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

        label = "Movies" if self.batch_type_var.get() == "movie" else "TV Series"
        skip_note = ""
        if self.skip_existing_var.get():
            skip_note = ("\n\nFolders that already contain poster.jpg + .nfo "
                         "will be skipped.")

        if not messagebox.askyesno(
            "Add subfolders",
            f"Add {len(subs)} subfolders of:\n{parent}\n\n"
            f"They will be added as: {label}.{skip_note}"):
            return

        added = 0
        skipped_complete = 0
        skipped_dupe = 0

        for p in subs:
            if any(r["path"] == p for r in self.row_data):
                skipped_dupe += 1
                continue
            if self.skip_existing_var.get() and has_complete_metadata(p):
                self._log(f"⏭️ Skipped (already has poster.jpg + .nfo): "
                          f"{p.name}")
                skipped_complete += 1
                continue
            if self._add_folder_path(p, silent=True):
                added += 1

        parts = [f"Added {added} of {len(subs)} subfolders."]
        if skipped_complete:
            parts.append(f"Skipped {skipped_complete} already-complete.")
        if skipped_dupe:
            parts.append(f"Skipped {skipped_dupe} already-in-list.")
        self._log(" ".join(parts))

        if skipped_complete:
            messagebox.showinfo(
                "Added",
                " ".join(parts) + "\n\n"
                "Folders with existing poster.jpg + .nfo were NOT added.\n"
                "Untick 'Skip folders that already have poster.jpg + .nfo' "
                "to include them.",
            )

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
            mtype = row.get("media_type", "movie")

            self.msg_queue.put({"type": "row", "index": i,
                                "fields": {"status": "Searching…"}})
            self.msg_queue.put({"type": "log",
                                "text": f"→ [{mtype}] {row['path'].name}  "
                                        f"(query: {title!r} year: {year!r})"})

            hit, used_v, used_y = search_robust(
                title, year, mtype,
                settings["api_key"], settings["language"],
                settings["include_adult"],
                log_cb=lambda t: self.msg_queue.put(
                    {"type": "log", "text": t}),
            )

            if hit:
                if mtype == "tv":
                    year_r = (hit.get("first_air_date") or "")[:4]
                    title_r = hit.get("name") or hit.get("original_name") or "?"
                else:
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
                                    "text": f"? No match: {row['path'].name}"})
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

            mtype = row.get("media_type", "movie")
            try:
                details = tmdb_details(int(tmdb_id), mtype,
                                       settings["api_key"],
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
                nfo = write_nfo(details, row["path"], mtype,
                                settings["actors"])
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
    # Queue polling
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
        elif "error" in status.lower() or status.startswith("❌"):
            tag = "error"
        elif status in ("No match",) or status.startswith("Query edited") \
                or status.startswith("Type ") or status.startswith("Type set"):
            tag = "warn"
        else:
            tag = ""

        self.tree.item(
            row["iid"], tags=(tag,) if tag else (),
            values=(
                i + 1,
                row["path"].name,
                row.get("media_type", "movie").upper(),
                row.get("hint", ""),
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
    # Context menu
    # ==================================================================

    def _on_right_click(self, event) -> None:
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _on_double_click(self, event) -> None:
        col = self.tree.identify_column(event.x)
        if col in ("#5", "#6"):
            self._on_edit_query()
        elif col == "#8":
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

    def _on_toggle_media_type(self) -> None:
        i, row = self._selected_row()
        if row is None:
            return
        new_type = "tv" if row.get("media_type") == "movie" else "movie"
        row["media_type"] = new_type
        row["match"] = ""
        row["tmdb_id"] = ""
        row["poster"] = ""
        row["nfo"] = ""
        row["status"] = f"Type toggled to {new_type.upper()} — re-search to apply"
        self._apply_row_update(i, {})
        self._log(f"{row['path'].name}: media_type → {new_type}")

    def _on_edit_query(self) -> None:
        i, row = self._selected_row()
        if row is None:
            return

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

        ttk.Label(frm, text="Media type:").grid(row=1, column=0, sticky="w")
        type_var = tk.StringVar(value=row.get("media_type", "movie"))
        ttk.Combobox(frm, textvariable=type_var,
                     values=["movie", "tv"], width=10,
                     state="readonly").grid(row=1, column=1, sticky="w",
                                            pady=(0, 6))

        ttk.Label(frm, text="Search title:").grid(row=2, column=0, sticky="w")
        title_var = tk.StringVar(value=row.get("query_title", ""))
        ttk.Entry(frm, textvariable=title_var, width=42).grid(
            row=2, column=1, sticky="ew", columnspan=2, pady=(0, 6))

        ttk.Label(frm, text="Year (optional):").grid(row=3, column=0, sticky="w")
        year_var = tk.StringVar(value=row.get("query_year", ""))
        ttk.Entry(frm, textvariable=year_var, width=10).grid(
            row=3, column=1, sticky="w", pady=(0, 6))

        auto_search_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Search this row after saving",
                        variable=auto_search_var).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))

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
            row["media_type"]  = type_var.get()
            row["match"] = ""
            row["tmdb_id"] = ""
            row["poster"] = ""
            row["nfo"] = ""
            row["status"] = "Query edited — re-search to apply"
            self._apply_row_update(i, {})
            self._log(f"Query for {row['path'].name}: "
                      f"title={new_title!r} year={new_year!r} "
                      f"type={row['media_type']}")
            dlg.destroy()

            if auto_search_var.get() and new_title:
                self._reprocess_row_async(i)

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Cancel",
                   command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Save", command=_ok).pack(side=tk.RIGHT,
                                                        padx=(0, 6))

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
            mtype = row.get("media_type", "movie")
            tmdb_id = row.get("override_id") or row.get("tmdb_id")

            if not tmdb_id:
                title = row.get("query_title") or ""
                year  = row.get("query_year") or None
                self.msg_queue.put({"type": "row", "index": i,
                                    "fields": {"status": "Searching…"}})
                self.msg_queue.put({"type": "log",
                                    "text": f"→ [{mtype}] {row['path'].name}"})
                hit, _, _ = search_robust(
                    title, year, mtype,
                    settings["api_key"], settings["language"],
                    settings["include_adult"],
                    log_cb=lambda t: self.msg_queue.put(
                        {"type": "log", "text": t}))
                if not hit:
                    self.msg_queue.put({"type": "row", "index": i,
                                        "fields": {"status": "No match"}})
                    return
                tmdb_id = hit["id"]
                if mtype == "tv":
                    year_r = (hit.get("first_air_date") or "")[:4]
                    title_r = hit.get("name") or hit.get("original_name") or "?"
                else:
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

            details = tmdb_details(int(tmdb_id), mtype,
                                   settings["api_key"],
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
                nfo = write_nfo(details, row["path"], mtype,
                                settings["actors"])
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