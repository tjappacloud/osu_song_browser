import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
import pygame
import sys
import re
import time
import json

# try to import Pillow for image thumbnails
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    Image = None
    ImageTk = None
    HAS_PIL = False

# try to import mutagen for reading mp3 tags
try:
    from mutagen._file import File as MutagenFile
    HAS_MUTAGEN = True
except Exception:
    MutagenFile = None
    HAS_MUTAGEN = False

# Supported audio extensions
SUPPORTED_AUDIO_EXTS = ('.mp3', '.ogg')
MIN_DURATION_SECONDS = 30
CACHE_FILENAME = '.osu_mp3_browser_cache.json'


def get_default_osu_songs_dir():
    # Typical osu! songs path on Windows
    home = Path.home()
    default = home / "AppData" / "Local" / "osu!" / "Songs"
    return default


def strip_leading_numbers(s: str) -> str:
    # Remove leading numeric IDs and separators (e.g. '311328 Foo' -> 'Foo')
    if not s:
        return s
    return re.sub(r'^\s*\d+[\s._-]*', '', s)


class OsuMP3Browser(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("osu! MP3 Browser")
        width = self.winfo_screenwidth()
        height = self.winfo_screenheight()
        # Start maximized (zoomed) on Windows; fallback to fullscreen-sized window
        try:
            self.state('zoomed')
        except Exception:
            self.geometry("%dx%d" % (width, height))

        # Allow toggling zoom with Escape
        self.bind('<Escape>', lambda e: self.toggle_fullscreen())

        # Initialize pygame mixer
        try:
            pygame.mixer.init()
        except Exception as e:
            messagebox.showwarning("Audio init failed", f"pygame.mixer.init() failed: {e}")
        # default volume (0.0 - 1.0)
        self.volume_var = tk.DoubleVar(value=0.8)
        try:
            if pygame.mixer.get_init():
                pygame.mixer.music.set_volume(self.volume_var.get())
        except Exception:
            pass

        # minimum duration (seconds) configurable via UI
        self.min_duration_var = tk.IntVar(value=MIN_DURATION_SECONDS)
        # dark mode toggle
        self.dark_mode_var = tk.BooleanVar(value=False)

        self.songs_dir = get_default_osu_songs_dir()
        # diagnostic: print songs_dir info
        try:
            print(f"Osu songs dir: {self.songs_dir} (exists={self.songs_dir.exists()})")
            if self.songs_dir.exists():
                try:
                    children = list(self.songs_dir.iterdir())[:10]
                    print(f"Top entries in songs dir: {[p.name for p in children]}")
                except Exception as _:
                    pass
        except Exception:
            pass
        # store tuples of (Path, folder_title) where folder_title is the parent folder name
        self.all_mp3_paths = []
        self.mp3_paths = []  # list of (Path, display_title)
        # quick membership set of known paths to avoid duplicates during incremental scans
        self._seen_paths = set()

        # UI
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=6)

        self.dir_label = ttk.Label(top, text=f"Songs dir: {self.songs_dir}")
        self.dir_label.pack(side=tk.LEFT, expand=True)

        browse_btn = ttk.Button(top, text="Browse...", command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT)
        # Manual scan button for debugging/refresh
        scan_btn = ttk.Button(top, text="Scan Now", command=lambda: threading.Thread(target=self.scan_and_populate, daemon=True).start())
        scan_btn.pack(side=tk.RIGHT, padx=(6, 0))
        # Dark mode toggle
        try:
            self.dark_check = ttk.Checkbutton(top, text="Dark Mode", variable=self.dark_mode_var, command=self._on_theme_changed)
            self.dark_check.pack(side=tk.RIGHT, padx=(6, 0))
        except Exception:
            try:
                self.dark_check = tk.Checkbutton(top, text="Dark Mode", variable=self.dark_mode_var, command=self._on_theme_changed)
                self.dark_check.pack(side=tk.RIGHT, padx=(6, 0))
            except Exception:
                pass
        # Search entry
        search_frame = ttk.Frame(self)
        search_frame.pack(fill=tk.X, padx=8)
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.search_entry.bind('<KeyRelease>', lambda e: self.refresh_list())
        clear_btn = ttk.Button(search_frame, text="Clear", command=self._clear_search)
        clear_btn.pack(side=tk.LEFT, padx=6)

        mid = ttk.Frame(self)
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        left = ttk.Frame(mid)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(left, activestyle='none')
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind('<Double-1>', self.on_double_click)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)

        scrollbar = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        right = ttk.Frame(mid, width=240)
        right.pack(side=tk.RIGHT, fill=tk.Y)
        # Background thumbnail (will be filled on selection)
        self.meta_image_label = ttk.Label(right)
        self.meta_image_label.pack(anchor=tk.CENTER, padx=6, pady=6)

        # Metadata labels
        self.meta_title = ttk.Label(right, text="Title: ")
        self.meta_title.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_artist = ttk.Label(right, text="Artist: ")
        self.meta_artist.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_album = ttk.Label(right, text="Album: ")
        self.meta_album.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_duration = ttk.Label(right, text="Duration: ")
        self.meta_duration.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_path = ttk.Label(right, text="Path: ", wraplength=220)
        self.meta_path.pack(anchor=tk.W, padx=6, pady=4)

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, padx=8, pady=6)

        # Now playing area (shows thumbnail and song title) - placed just above controls
        now_frame = ttk.Frame(self)
        now_frame.pack(fill=tk.X, padx=8, pady=(0, 4))
        self.now_image_label = ttk.Label(now_frame)
        self.now_image_label.pack(side=tk.LEFT, padx=(0, 8))
        now_right = ttk.Frame(now_frame)
        now_right.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.now_title_label = ttk.Label(now_right, text="Now: Not playing")
        self.now_title_label.pack(anchor=tk.W)
        # progress bar and time label
        # use a finer-grained internal scale (0-1000) for smoother progress updates
        self.progress = ttk.Progressbar(now_right, orient=tk.HORIZONTAL, mode='determinate', length=400, maximum=1000)
        self.progress.pack(fill=tk.X, pady=(4, 0))
        self.time_label = ttk.Label(now_right, text="0:00 / 0:00")
        self.time_label.pack(anchor=tk.W)
        # playback tracking
        self._playing_path = None
        self._progress_after_id = None
        # manual timing for smoother progress and seeking
        self._start_time = None
        self._pause_time = None
        self._paused_offset = 0.0
        # bind progress seeking events
        try:
            self.progress.bind('<Button-1>', self.on_progress_click)
            self.progress.bind('<B1-Motion>', self.on_progress_click)
        except Exception:
            pass

        self.play_btn = ttk.Button(bottom, text="Play", command=self.play_selected)
        self.play_btn.pack(side=tk.LEFT)

        self.pause_btn = ttk.Button(bottom, text="Pause", command=self.toggle_pause)
        self.pause_btn.pack(side=tk.LEFT, padx=6)

        self.stop_btn = ttk.Button(bottom, text="Stop", command=self.stop)
        self.stop_btn.pack(side=tk.LEFT)

        # Volume control
        self.volume_label = ttk.Label(bottom, text=f"Vol: {int(self.volume_var.get()*100)}%")
        self.volume_label.pack(side=tk.LEFT, padx=(8, 4))
        # Use a ttk.Scale for volume (0.0 - 1.0)
        self.volume_scale = ttk.Scale(bottom, from_=0.0, to=1.0, orient=tk.HORIZONTAL,
                          length=120, variable=self.volume_var,
                          command=self.on_volume_change)
        self.volume_scale.pack(side=tk.LEFT)

        # Minimum duration spinbox
        try:
            # ttk.Spinbox is available in newer tkinter versions
            self.min_label = ttk.Label(bottom, text="Min length (s):")
            self.min_label.pack(side=tk.LEFT, padx=(8, 4))
            self.min_spin = ttk.Spinbox(bottom, from_=0, to=3600, textvariable=self.min_duration_var, width=6, command=lambda: threading.Thread(target=self._on_min_duration_changed, daemon=True).start())
            self.min_spin.pack(side=tk.LEFT)
        except Exception:
            # fallback to tk.Spinbox
            try:
                self.min_label = ttk.Label(bottom, text="Min length (s):")
                self.min_label.pack(side=tk.LEFT, padx=(8, 4))
                self.min_spin = tk.Spinbox(bottom, from_=0, to=3600, textvariable=self.min_duration_var, width=6, command=lambda: threading.Thread(target=self._on_min_duration_changed, daemon=True).start())
                self.min_spin.pack(side=tk.LEFT)
            except Exception:
                pass

        self.current_label = ttk.Label(bottom, text="Not playing")
        self.current_label.pack(side=tk.RIGHT)

        # scan on start (in background)
        self.after(100, lambda: threading.Thread(target=self.scan_and_populate, daemon=True).start())

        self.paused = False
        # metadata cache: path -> dict
        self._metadata = {}
        # counter for excluded short files during scanning (updated on main thread)
        self._excluded_short = 0
        # persistent cache file path
        try:
            self.cache_path = Path.home() / CACHE_FILENAME
        except Exception:
            self.cache_path = Path(CACHE_FILENAME)

        # try to load existing cache so UI can populate faster (also loads theme)
        try:
            self._load_cache()
            # apply theme from cache before showing UI
            try:
                self.apply_theme()
            except Exception:
                pass
            # apply cached entries to UI immediately
            try:
                self.after(0, self._apply_cache_to_ui)
            except Exception:
                pass
        except Exception:
            pass

    def _begin_scan_ui(self):
        # Clear current visible lists and show scanning state (must run on main thread)
        try:
            self.listbox.delete(0, tk.END)
        except Exception:
            pass
        try:
            self.mp3_paths.clear()
        except Exception:
            pass
        try:
            self.all_mp3_paths.clear()
        except Exception:
            pass
        self._excluded_short = 0
        try:
            self.current_label.config(text="Scanning...")
        except Exception:
            pass

    def _apply_cache_to_ui(self):
        """Populate the visible list from the loaded cache quickly (main thread)."""
        try:
            # Clear current visible lists
            try:
                self.listbox.delete(0, tk.END)
            except Exception:
                pass
            self.mp3_paths.clear()
            for path, folder_title in self.all_mp3_paths:
                # populate seen set from cache so future scans don't duplicate
                try:
                    self._seen_paths.add(str(path))
                except Exception:
                    pass
                # apply current search filter
                q = (self.search_var.get() or '').strip().lower()
                if q:
                    meta = self._metadata.get(str(path), {})
                    searchable = [folder_title.lower(), str(meta.get('title') or '').lower(), str(meta.get('artist') or '').lower()]
                    if not any(q in s for s in searchable):
                        continue
                self.mp3_paths.append((path, folder_title))
                try:
                    self.listbox.insert(tk.END, folder_title)
                except Exception:
                    pass
            try:
                self.current_label.config(text=f"Found {len(self.all_mp3_paths)} audio files (cached)")
            except Exception:
                pass
        except Exception:
            pass

    def apply_theme(self):
        """Apply the chosen theme (dark/light) to the UI widgets and ttk styles."""
        try:
            dark = bool(self.dark_mode_var.get())
            style = ttk.Style()
            # prefer 'clam' theme for better style control where available
            try:
                style.theme_use('clam')
            except Exception:
                try:
                    style.theme_use('default')
                except Exception:
                    pass

            if dark:
                # explicit dark palette
                bg = '#2e2e2e'
                fg = '#eaeaea'
                entry_bg = '#3a3a3a'
                list_bg = '#1e1e1e'
                select_bg = '#555555'
                button_bg = '#3a3a3a'
            else:
                # explicit light palette (avoid None to prevent type issues)
                bg = '#f0f0f0'
                fg = '#000000'
                entry_bg = '#ffffff'
                list_bg = '#ffffff'
                select_bg = '#3399ff'
                button_bg = '#e0e0e0'

            # configure ttk styles
            try:
                style.configure('TFrame', background=bg)
                style.configure('TLabel', background=bg, foreground=fg)
                style.configure('TButton', background=button_bg, foreground=fg)
                style.configure('TEntry', fieldbackground=entry_bg, foreground=fg)
                style.configure('Horizontal.TScale', background=bg)
                style.configure('TScrollbar', background=bg)
                # progressbar styling (may vary by platform)
                try:
                    style.configure('TProgressbar', troughcolor=bg, background=button_bg)
                except Exception:
                    pass
            except Exception:
                pass

            # apply to some direct tk widgets
            try:
                self.configure(bg=bg)
            except Exception:
                pass
            try:
                self.listbox.config(bg=list_bg, fg=fg, selectbackground=select_bg, highlightbackground=bg)
            except Exception:
                pass
        except Exception:
            pass

    def _on_theme_changed(self):
        """Callback when theme checkbox toggled: apply theme and save settings to cache."""
        try:
            self.apply_theme()
        except Exception:
            pass
        try:
            # save settings into cache immediately
            self._save_cache()
        except Exception:
            pass

    def _load_cache(self):
        """Load cached discovery file if present and validate entries.
        Cache format: list of {path, folder_title, meta: {...}} where meta may contain '__mtime' and '__size'.
        """
        try:
            if not getattr(self, 'cache_path', None):
                return
            if not self.cache_path.exists():
                return
            try:
                with self.cache_path.open('r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception:
                return

            # support both old-list format and new dict-with-settings format
            settings = {}
            if isinstance(data, dict):
                settings = data.get('settings', {}) or {}
                items = data.get('items') or data.get('out') or []
            else:
                items = data
            # apply settings (e.g., dark mode)
            try:
                if settings.get('dark_mode') is not None:
                    self.dark_mode_var.set(bool(settings.get('dark_mode')))
            except Exception:
                pass

            # validate and load
            self.all_mp3_paths.clear()
            for rec in items:
                try:
                    p = Path(rec.get('path', ''))
                    if not p.exists():
                        continue
                    st = p.stat()
                    mtime = int(st.st_mtime)
                    size = st.st_size
                    meta = rec.get('meta', {}) or {}
                    rec_mtime = int(meta.get('__mtime') or rec.get('mtime') or 0)
                    rec_size = int(meta.get('__size') or rec.get('size') or 0)
                    if rec_mtime and rec_size and rec_mtime == mtime and rec_size == size:
                        folder_title = rec.get('folder_title') or strip_leading_numbers(p.parent.name)
                        self.all_mp3_paths.append((p, folder_title))
                        # keep metadata including internal markers
                        meta['__mtime'] = mtime
                        meta['__size'] = size
                        self._metadata[str(p)] = dict(meta)
                        try:
                            self._seen_paths.add(str(p))
                        except Exception:
                            pass
                except Exception:
                    continue
        except Exception:
            return

    def _save_cache(self):
        """Persist current discovery results to cache for faster next startup."""
        try:
            if not getattr(self, 'cache_path', None):
                return
            out = []
            for p, folder_title in self.all_mp3_paths:
                try:
                    key = str(p)
                    meta = dict(self._metadata.get(key, {}))
                    # ensure mtime/size stored
                    try:
                        st = p.stat()
                        meta['__mtime'] = int(st.st_mtime)
                        meta['__size'] = st.st_size
                    except Exception:
                        pass
                    out.append({'path': key, 'folder_title': folder_title, 'meta': meta})
                except Exception:
                    continue
            try:
                payload = {'items': out, 'settings': {'dark_mode': bool(self.dark_mode_var.get())}}
                with self.cache_path.open('w', encoding='utf-8') as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        except Exception:
            pass

    def _inc_excluded_short(self):
        try:
            self._excluded_short += 1
            # update status label
            try:
                min_d = self.min_duration_var.get() if hasattr(self, 'min_duration_var') else MIN_DURATION_SECONDS
                self.current_label.config(text=f"Found {len(self.all_mp3_paths)} audio files (excluded {self._excluded_short} < {min_d}s)")
            except Exception:
                pass
        except Exception:
            pass

    def _on_min_duration_changed(self):
        """Called when the min duration spinbox changes: trigger a re-scan so UI reflects new cutoff."""
        try:
            # kick off a background re-scan (scan_and_populate already schedules UI updates)
            threading.Thread(target=self.scan_and_populate, daemon=True).start()
        except Exception:
            pass

    def _add_discovered_file(self, full: Path, folder_title: str, meta: dict):
        """Add a single discovered file to internal lists and the visible listbox (main thread)."""
        try:
            key = str(full)
            # avoid adding duplicates if this path was already known/displayed
            if key in self._seen_paths:
                # still merge metadata if provided
                try:
                    if meta:
                        self._metadata[key] = {**self._metadata.get(key, {}), **meta}
                except Exception:
                    pass
                return

            # Re-check duration here to avoid adding files that were mis-measured
            try:
                dur = meta.get('duration') if meta else 0
            except Exception:
                dur = 0
            if not dur:
                try:
                    dur = self.ensure_duration(full)
                except Exception:
                    dur = 0
            min_d = self.min_duration_var.get() if hasattr(self, 'min_duration_var') else MIN_DURATION_SECONDS
            if dur and dur < min_d:
                # count as excluded and do not add
                try:
                    self._inc_excluded_short()
                except Exception:
                    pass
                return
            key = str(full)
            # mark as seen so future scans won't re-add
            try:
                self._seen_paths.add(key)
            except Exception:
                pass
            # merge metadata for this file
            try:
                if meta:
                    self._metadata[key] = meta
            except Exception:
                pass
            # add to master list
            try:
                self.all_mp3_paths.append((full, folder_title))
            except Exception:
                pass

            # decide if matches current search
            q = (self.search_var.get() or '').strip().lower()
            match = True
            if q:
                searchable = [folder_title.lower()]
                try:
                    if meta.get('title'):
                        searchable.append(str(meta.get('title')).lower())
                except Exception:
                    pass
                try:
                    if meta.get('artist'):
                        searchable.append(str(meta.get('artist')).lower())
                except Exception:
                    pass
                match = any(q in s for s in searchable)

            if match:
                # add to visible list
                try:
                    self.mp3_paths.append((full, folder_title))
                    self.listbox.insert(tk.END, folder_title)
                except Exception:
                    pass

            # update status label with running count
            try:
                if self._excluded_short:
                    self.current_label.config(text=f"Found {len(self.all_mp3_paths)} audio files (excluded {self._excluded_short} < {MIN_DURATION_SECONDS}s)")
                else:
                    self.current_label.config(text=f"Found {len(self.all_mp3_paths)} audio files")
            except Exception:
                pass
        except Exception:
            pass

    def browse_folder(self):
        path = filedialog.askdirectory(initialdir=str(self.songs_dir) if self.songs_dir.exists() else None)
        if path:
            self.songs_dir = Path(path)
            self.dir_label.config(text=f"Songs dir: {self.songs_dir}")
            threading.Thread(target=self.scan_and_populate, daemon=True).start()

    def scan_and_populate(self):
        # Perform file discovery and metadata retrieval on background thread,
        # but apply UI updates on the main thread to avoid tkinter thread-safety issues.
        try:
            if not self.songs_dir.exists():
                # schedule UI update to show not found
                self.after(0, lambda: self.listbox.insert(tk.END, "(Songs directory not found)"))
                return
        except Exception as e:
            print(f"Error checking songs_dir: {e}")
            self.after(0, lambda: self.listbox.insert(tk.END, "(Songs directory error)"))
            return

        # indicate scanning but do not clear the currently-displayed list;
        # we want incremental discovery that preserves what's already shown
        try:
            self.after(0, lambda: self.current_label.config(text="Scanning..."))
        except Exception:
            pass

        local_all = []
        local_meta = {}
        excluded_short = 0

        # Process each folder and pick only the first supported audio file in it
        for root, dirs, files in sorted(os_walk(self.songs_dir)):
            try:
                # find the first filename in sorted order that matches supported extensions
                first_fn = None
                for fn in sorted(files):
                    if fn.lower().endswith(SUPPORTED_AUDIO_EXTS):
                        first_fn = fn
                        break
                if not first_fn:
                    continue
                full = Path(root) / first_fn
                # try to reuse cached metadata if file unchanged
                key = str(full)
                meta = {}
                try:
                    st = full.stat()
                    mtime = int(st.st_mtime)
                    size = st.st_size
                except Exception:
                    mtime = None
                    size = None

                cached = self._metadata.get(key)
                if cached and mtime is not None and size is not None and cached.get('__mtime') == mtime and cached.get('__size') == size:
                    # reuse cached metadata (strip internal keys when later used)
                    meta = {k: v for k, v in cached.items() if not k.startswith('__')}
                    # ensure we copy the internal markers into local_meta so save later keeps them
                    local_meta[key] = dict(cached)
                else:
                    # read metadata (may be slow) into local dict
                    meta = get_mp3_metadata(full) if HAS_MUTAGEN else {}
                    # attach mtime/size markers for caching
                    try:
                        if mtime is not None:
                            meta['__mtime'] = mtime
                        if size is not None:
                            meta['__size'] = size
                    except Exception:
                        pass
                    local_meta[key] = dict(meta)
                # ensure duration is known (may compute and cache). Prefer cached value.
                dur = meta.get('duration') or 0
                if not dur:
                    try:
                        dur = self.ensure_duration(full)
                        # store duration into local_meta for saving
                        try:
                            local_meta[str(full)]['duration'] = dur
                        except Exception:
                            pass
                    except Exception:
                        dur = 0
                # skip very short files (count as excluded)
                min_d = self.min_duration_var.get() if hasattr(self, 'min_duration_var') else MIN_DURATION_SECONDS
                if dur and dur < min_d:
                    excluded_short += 1
                    try:
                        self.after(0, self._inc_excluded_short)
                    except Exception:
                        pass
                    continue

                folder_title = strip_leading_numbers(full.parent.name)
                local_all.append((full, folder_title))
                # add this file to the UI immediately
                try:
                    self.after(0, lambda f=full, t=folder_title, m=meta: self._add_discovered_file(f, t, m))
                except Exception:
                    pass
            except Exception:
                # ignore errors per-folder
                continue

        # Apply results to UI on main thread
        def apply_results():
            try:
                # merge remaining metadata
                try:
                    self._metadata.update(local_meta)
                except Exception:
                    pass
                # final status update
                count = len(self.all_mp3_paths)
                if count == 0:
                    try:
                        self.listbox.insert(tk.END, "(No audio files found in Songs directory)")
                    except Exception:
                        pass
                if excluded_short:
                    try:
                        min_d = self.min_duration_var.get() if hasattr(self, 'min_duration_var') else MIN_DURATION_SECONDS
                        self.current_label.config(text=f"Found {count} audio files (excluded {excluded_short} < {min_d}s)")
                    except Exception:
                        pass
                else:
                    try:
                        self.current_label.config(text=f"Found {count} audio files")
                    except Exception:
                        pass
                print(f"scan_and_populate: found {count} audio files in {self.songs_dir} (excluded_short={excluded_short})")
                try:
                    # persist cache for faster startup next time
                    self._save_cache()
                except Exception:
                    pass
            except Exception as e:
                print(f"Error applying scan results: {e}")

        self.after(0, apply_results)

    def play_selected(self):
        idx = self.listbox.curselection()
        if not idx:
            messagebox.showinfo("Select", "Please select an audio file from the list.")
            return
        index = idx[0]
        try:
            path = self.mp3_paths[index][0]
        except IndexError:
            return
        self._play_path(path)

    def on_double_click(self, event):
        self.play_selected()

    def _play_path(self, path: Path):
        try:
            pygame.mixer.music.stop()
            # load and play in background
            pygame.mixer.music.load(str(path))
            pygame.mixer.music.play()
            # display folder title as the song name
            # find matching entry in mp3_paths to get the folder title
            folder_title = strip_leading_numbers(path.parent.name)
            # if we stored folder title in mp3_paths, prefer that
            for p, t in self.mp3_paths:
                if p == path:
                    folder_title = t
                    break
            self.current_label.config(text=f"Playing: {folder_title}")
            # Update now-playing display (thumbnail + title)
            self.now_title_label.config(text=f"Now: {folder_title}")
            # start updating progress
            self._playing_path = path
            # Initialize manual timing base so progress/time are consistent
            try:
                self._start_time = time.time()
            except Exception:
                self._start_time = None
            self._pause_time = None
            self._paused_offset = 0.0
            # cancel previous updater if any
            if self._progress_after_id:
                try:
                    self.after_cancel(self._progress_after_id)
                except Exception:
                    pass
                self._progress_after_id = None
            # ensure pause button shows correct action when starting playback
            try:
                self.pause_btn.config(text="Pause")
            except Exception:
                pass
            self.update_progress()
            # load background thumbnail if available
            bg = get_osu_background(path.parent)
            if bg and HAS_PIL:
                    try:
                        from PIL import Image as PILImage, ImageTk as PILImageTk
                        img = PILImage.open(bg)
                        # small thumbnail for now-playing (fit into 96x54)
                        resampling = getattr(PILImage, 'Resampling', None)
                        if resampling is not None:
                            resample = getattr(resampling, 'LANCZOS', None)
                        else:
                            resample = getattr(PILImage, 'LANCZOS', None)
                        if resample is not None:
                            img.thumbnail((96, 54), resample)
                        else:
                            img.thumbnail((96, 54))
                        photo = PILImageTk.PhotoImage(img)
                        self.now_image_label.config(image=photo)
                        setattr(self.now_image_label, '_photo_ref', photo)
                    except Exception:
                        self.now_image_label.config(image='')
                        if hasattr(self.now_image_label, '_photo_ref'):
                            delattr(self.now_image_label, '_photo_ref')
            else:
                self.now_image_label.config(image='')
                if hasattr(self.now_image_label, '_photo_ref'):
                    delattr(self.now_image_label, '_photo_ref')
            self.paused = False
        except Exception as e:
            messagebox.showerror("Playback error", f"Failed to play {path}: {e}")

    def toggle_pause(self):
        if not pygame.mixer.get_init():
            return
        # If nothing is playing, do nothing
        if not self._playing_path:
            return

        # Toggle paused state. Do not rely on pygame.mixer.music.get_busy()
        # because some backends may report False while paused.
        if not self.paused:
            # Try to pause via pygame; if it fails, we still set the manual pause time
            try:
                pygame.mixer.music.pause()
            except Exception:
                pass
            self.paused = True
            self.pause_btn.config(text="Resume")
            self.current_label.config(text=self.current_label.cget("text") + " (paused)")
            # record pause time for manual timing calculations
            try:
                self._pause_time = time.time()
            except Exception:
                self._pause_time = None
        else:
            # Attempt to unpause; if unpause isn't supported by backend, fall back
            # to restarting playback at the manual paused offset.
            unpaused = False
            try:
                pygame.mixer.music.unpause()
                unpaused = True
            except Exception:
                unpaused = False

            if not unpaused:
                # compute paused position from manual timer
                pos_sec = 0
                try:
                    if self._start_time is not None and self._pause_time is not None:
                        pos_sec = int(self._pause_time - self._start_time)
                except Exception:
                    pos_sec = 0
                try:
                    # seek_to will attempt best-effort methods to resume at pos_sec
                    self.seek_to(pos_sec)
                except Exception:
                    try:
                        # as final fallback, just play from start
                        pygame.mixer.music.play()
                    except Exception:
                        pass

            self.paused = False
            self.pause_btn.config(text="Pause")
            # remove (paused) suffix
            txt = self.current_label.cget("text").replace(" (paused)", "")
            self.current_label.config(text=txt)
            # adjust manual timing to account for pause duration
            try:
                if self._pause_time and self._start_time:
                    paused_duration = time.time() - self._pause_time
                    self._start_time += paused_duration
            except Exception:
                pass
            self._pause_time = None

    def stop(self):
        if pygame.mixer.get_init():
            pygame.mixer.music.stop()
        self.current_label.config(text="Not playing")
        # clear now-playing and cancel progress updates
        self._playing_path = None
        # clear manual timing
        self._start_time = None
        self._pause_time = None
        self._paused_offset = 0.0
        # reset pause button state
        try:
            self.pause_btn.config(text="Pause")
        except Exception:
            pass
        self.paused = False
        if self._progress_after_id:
            try:
                self.after_cancel(self._progress_after_id)
            except Exception:
                pass
            self._progress_after_id = None
        self.now_title_label.config(text="Now: Not playing")
        self.now_image_label.config(image='')
        if hasattr(self.now_image_label, '_photo_ref'):
            delattr(self.now_image_label, '_photo_ref')
        self.progress['value'] = 0
        self.time_label.config(text="0:00 / 0:00")

    def toggle_fullscreen(self, event=None):
        try:
            # Toggle between zoomed (maximized) and normal windowed state
            if self.state() == 'zoomed':
                self.state('normal')
                w = int(self.winfo_screenwidth() * 0.8)
                h = int(self.winfo_screenheight() * 0.8)
                self.geometry(f"{w}x{h}")
            else:
                self.state('zoomed')
        except Exception:
            pass

    def on_volume_change(self, val):
        """Callback for volume scale. `val` is a string from the scale command."""
        try:
            v = float(val)
        except Exception:
            try:
                v = self.volume_var.get()
            except Exception:
                return
        try:
            if pygame.mixer.get_init():
                pygame.mixer.music.set_volume(v)
        except Exception:
            pass
        try:
            # update label
            self.volume_label.config(text=f"Vol: {int(v*100)}%")
        except Exception:
            pass

    def on_select(self, event):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        try:
            path = self.mp3_paths[idx][0]
        except IndexError:
            return
        meta = self._metadata.get(str(path), {})
        # display song name based on folder name
        title = strip_leading_numbers(path.parent.name)
        artist = meta.get('artist') or ''
        album = meta.get('album') or ''
        duration = format_duration(meta.get('duration')) if meta.get('duration') else ''
        self.meta_title.config(text=f"Title: {title}")
        self.meta_artist.config(text=f"Artist: {artist}")
        self.meta_album.config(text=f"Album: {album}")
        self.meta_duration.config(text=f"Duration: {duration}")
        self.meta_path.config(text=f"Path: {path}")
        # Try to load background from the first .osu file in the folder
        bg = get_osu_background(path.parent)
        if bg and HAS_PIL:
                try:
                    from PIL import Image as PILImage, ImageTk as PILImageTk
                    img = PILImage.open(bg)
                    # create thumbnail keeping aspect ratio, fit into 220x140
                    # Pillow uses Image.Resampling.LANCZOS in newer versions; fall back if missing
                    resampling = getattr(PILImage, 'Resampling', None)
                    if resampling is not None:
                        resample = getattr(resampling, 'LANCZOS', None)
                    else:
                        resample = getattr(PILImage, 'LANCZOS', None)
                    if resample is not None:
                        img.thumbnail((220, 140), resample)
                    else:
                        img.thumbnail((220, 140))
                    photo = PILImageTk.PhotoImage(img)
                    self.meta_image_label.config(image=photo)
                    # retain reference on the label widget to avoid GC
                    setattr(self.meta_image_label, '_photo_ref', photo)
                except Exception:
                    # clear image on error
                    self.meta_image_label.config(image='')
                    if hasattr(self.meta_image_label, '_photo_ref'):
                        delattr(self.meta_image_label, '_photo_ref')
        else:
            # clear image if none found or PIL missing
            self.meta_image_label.config(image='')
            if hasattr(self.meta_image_label, '_photo_ref'):
                delattr(self.meta_image_label, '_photo_ref')

    def _clear_search(self):
        self.search_var.set('')
        self.refresh_list()

    def update_progress(self):
        """Poll playback position and update the progress bar and time label."""
        try:
            path = self._playing_path
            if not path or not pygame.mixer.get_init():
                return

            total = self._metadata.get(str(path), {}).get('duration') or 0
            # if duration unknown, try to compute and cache it
            if not total:
                total = self.ensure_duration(path)

            # Prefer manual timing base (self._start_time) for progress display so
            # clicking and seeking doesn't cause the UI to snap to 0 due to
            # backend get_pos() resets. Fall back to pygame.get_pos() if manual
            # timing isn't available.
            busy = pygame.mixer.music.get_busy()
            if not busy and not self.paused:
                # mark complete
                if total:
                    self.progress['value'] = 1000
                    self.time_label.config(text=f"{format_duration(total)} / {format_duration(total)}")
                # cancel further updates
                self._playing_path = None
                return

            # Compute position using manual base when possible
            pos_sec = 0
            try:
                if self._start_time is not None:
                    if self.paused and self._pause_time is not None:
                        pos_sec = int(self._pause_time - self._start_time)
                    elif not self.paused:
                        pos_sec = int(time.time() - self._start_time)
                    else:
                        pos_sec = 0
                else:
                    pos_ms = pygame.mixer.music.get_pos()
                    if pos_ms is None or pos_ms < 0:
                        pos_sec = 0
                    else:
                        pos_sec = int(pos_ms / 1000)
            except Exception:
                # fallback to pygame get_pos
                try:
                    pos_ms = pygame.mixer.music.get_pos()
                    pos_sec = int(pos_ms / 1000) if pos_ms and pos_ms >= 0 else 0
                except Exception:
                    pos_sec = 0

            if total:
                frac = min(1.0, pos_sec / total)
                self.progress['value'] = int(frac * 1000)
                self.time_label.config(text=f"{format_duration(pos_sec)} / {format_duration(total)}")
            else:
                # unknown total
                self.progress['value'] = 0
                self.time_label.config(text=f"{format_duration(pos_sec)} / 0:00")

            # schedule next poll
            self._progress_after_id = self.after(500, self.update_progress)
        except Exception:
            self._progress_after_id = None

    def refresh_list(self):
        """Refresh visible listbox entries based on `self.search_var`.
        Matches against folder title, cached tag title, and artist (case-insensitive substring).
        """
        q = (self.search_var.get() or '').strip().lower()
        self.listbox.delete(0, tk.END)
        self.mp3_paths.clear()
        for path, folder_title in self.all_mp3_paths:
            # gather searchable strings
            searchable = [folder_title.lower()]
            meta = self._metadata.get(str(path), {})
            if meta.get('title'):
                searchable.append(str(meta.get('title')).lower())
            if meta.get('artist'):
                searchable.append(str(meta.get('artist')).lower())

            # decide if item matches query
            match = True
            if q:
                match = any(q in s for s in searchable)

            if match:
                self.mp3_paths.append((path, folder_title))
                self.listbox.insert(tk.END, folder_title)

    def on_progress_click(self, event):
        """Handle click/drag on the progress bar to seek."""
        try:
            widget = event.widget
            w = widget.winfo_width()
            if w <= 0:
                return
            x = event.x
            frac = max(0.0, min(1.0, x / w))
            # compute target seconds
            if not self._playing_path:
                return
            total = self._metadata.get(str(self._playing_path), {}).get('duration') or self.ensure_duration(self._playing_path)
            if not total:
                return
            target = frac * total
            self.seek_to(target)
        except Exception:
            pass

    def seek_to(self, pos_sec: float):
        """Seek to pos_sec (seconds) in the currently playing file."""
        if not self._playing_path:
            return
        # clamp
        total = self._metadata.get(str(self._playing_path), {}).get('duration') or self.ensure_duration(self._playing_path)
        if total and pos_sec > total:
            pos_sec = total
        try:
            # Attempt several seek methods in order for best compatibility.
            # 1) pygame.mixer.music.set_pos(pos) then play() — works on some backends.
            # 2) pygame.mixer.music.play(0, pos) — many versions support start parameter for MP3.
            # 3) fallback: restart playback from beginning.
            success = False
            try:
                pygame.mixer.music.stop()
            except Exception:
                pass

            # Try set_pos first
            try:
                pygame.mixer.music.set_pos(float(pos_sec))
                pygame.mixer.music.play()
                success = True
            except Exception:
                success = False

            if not success:
                try:
                    # try play with start position (some backends accept float)
                    pygame.mixer.music.play(0, float(pos_sec))
                    success = True
                except TypeError:
                    try:
                        pygame.mixer.music.play(0, pos_sec)
                        success = True
                    except Exception:
                        success = False
                except Exception:
                    success = False

            if not success:
                # final fallback: just play from start
                try:
                    pygame.mixer.music.play()
                except Exception:
                    pass
            # update manual timing regardless of which method succeeded
            self._start_time = time.time() - float(pos_sec)
            self._pause_time = None
            self._paused_offset = 0.0
            self.paused = False
            # restart progress polling
            if self._progress_after_id:
                try:
                    self.after_cancel(self._progress_after_id)
                except Exception:
                    pass
            self.update_progress()
        except Exception:
            pass

    def ensure_duration(self, path: Path) -> int:
        """Ensure we have a cached duration (seconds) for `path`. Tries Mutagen then pygame.mixer.Sound.
        Returns duration in seconds (int) or 0 if unknown.
        """
        key = str(path)
        meta = self._metadata.get(key, {})
        dur = meta.get('duration') or 0
        if dur:
            return dur

        # Try mutagen first
        if HAS_MUTAGEN and MutagenFile is not None:
            try:
                audio = MutagenFile(key)
                if audio and getattr(audio, 'info', None):
                    length = int(getattr(audio.info, 'length', 0) or 0)
                    if length:
                        meta['duration'] = length
                        self._metadata[key] = meta
                        return length
            except Exception:
                pass

        # Fall back to pygame.mixer.Sound (may use more memory)
        try:
            snd = pygame.mixer.Sound(key)
            length = int(snd.get_length() or 0)
            if length:
                meta['duration'] = length
                self._metadata[key] = meta
                return length
        except Exception:
            pass

        return 0


def os_walk(path: Path):
    # Simple wrapper so we can mock/test easily
    for root, dirs, files in __import__('os').walk(path):
        yield root, dirs, files


def get_mp3_metadata(path: Path) -> dict:
    """Return a small metadata dict for the mp3: title, artist, album, duration (seconds).
    Requires mutagen; returns empty dict if unavailable or on error.
    """
    if not HAS_MUTAGEN or MutagenFile is None:
        return {}

    try:
        # Try Easy interface first (maps common names like 'title', 'artist')
        audio_easy = MutagenFile(str(path), easy=True)
        meta = {}
        title = None
        artist = None
        album = None
        duration = None

        if audio_easy:
            try:
                title = audio_easy.get('title', [None])[0]
                artist = audio_easy.get('artist', [None])[0]
                album = audio_easy.get('album', [None])[0]
            except Exception:
                pass
            try:
                duration = int(getattr(audio_easy.info, 'length', 0))
            except Exception:
                duration = None

        # If we didn't get a title, try raw tags (ID3 frames) for TIT2
        if not title:
            audio_raw = MutagenFile(str(path))
            if audio_raw and getattr(audio_raw, 'tags', None) is not None:
                tags = audio_raw.tags
                # ID3 frames: TIT2 is title, TPE1 artist, TALB album
                try:
                    tit2 = tags.get('TIT2')
                    if tit2 is not None:
                        # frame may have .text
                        val = getattr(tit2, 'text', None)
                        if val:
                            title = val[0] if isinstance(val, (list, tuple)) else val
                except Exception:
                    pass
                try:
                    tpe1 = tags.get('TPE1')
                    if tpe1 is not None:
                        val = getattr(tpe1, 'text', None)
                        if val:
                            artist = val[0] if isinstance(val, (list, tuple)) else val
                except Exception:
                    pass
                try:
                    talb = tags.get('TALB')
                    if talb is not None:
                        val = getattr(talb, 'text', None)
                        if val:
                            album = val[0] if isinstance(val, (list, tuple)) else val
                except Exception:
                    pass
                try:
                    if duration is None:
                        duration = int(getattr(audio_raw.info, 'length', 0))
                except Exception:
                    duration = duration

        # Fallback: use filename stem with numbers stripped
        if not title:
            title = strip_leading_numbers(path.stem)

        if title:
            meta['title'] = title
        if artist:
            meta['artist'] = artist
        if album:
            meta['album'] = album
        if duration:
            meta['duration'] = duration

        return meta
    except Exception:
        return {}


def get_osu_background(folder: Path) -> Path | None:
    """Find the first .osu file in folder and parse its [Events] section for a background image.
    Returns the resolved Path to the image if found and exists, otherwise None.
    """
    try:
        # find first .osu file
        for p in sorted(folder.iterdir()):
            if p.suffix.lower() == '.osu':
                osu_path = p
                break
        else:
            return None

        with osu_path.open('r', encoding='utf-8', errors='ignore') as f:
            in_events = False
            for line in f:
                line = line.strip()
                if line == '[Events]':
                    in_events = True
                    continue
                if in_events:
                    if line.startswith('['):
                        # next section
                        break
                    # look for a quoted filename (common format: 0,0,"bg.jpg",0)
                    if '"' in line:
                        # extract first quoted string
                        m = re.search(r'"([^"]+)"', line)
                        if m:
                            imgname = m.group(1)
                            # check extension
                            if imgname.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                                candidate = folder / imgname
                                if candidate.exists():
                                    return candidate
                                # sometimes path may include subfolders
                                candidate2 = folder / imgname.replace('\\', '/')
                                if candidate2.exists():
                                    return candidate2
                    # some osu files may list backgrounds without quotes (rare)
            return None
    except Exception:
        return None


def format_duration(sec: int) -> str:
    if not sec:
        return ''
    m, s = divmod(int(sec), 60)
    return f"{m}:{s:02d}"


def main():
    app = OsuMP3Browser()
    app.mainloop()


if __name__ == '__main__':
    main()
