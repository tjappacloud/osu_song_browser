"""Main UI class for osu! Song Browser."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
try:
    import tkinter.font as tkfont
except Exception:
    tkfont = None
import threading
import time
import json
import random
from pathlib import Path

from .config import get_default_osu_songs_dir, SUPPORTED_AUDIO_EXTS, MIN_DURATION_SECONDS, CACHE_FILENAME
from .utils import strip_leading_numbers, parse_artist_from_folder, format_duration, os_walk
from .metadata import get_mp3_metadata, get_osu_background, ensure_duration
from . import audio
from .playlist import PlaylistStore

# try to import Pillow for image thumbnails
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    Image = None
    ImageTk = None
    HAS_PIL = False


class OsuMP3Browser(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("osu! Song Browser")
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
        if not audio.init_audio():
            messagebox.showwarning("Audio init failed", "pygame.mixer.init() failed")
        
        # default volume (0.0 - 1.0)
        self.volume_var = tk.DoubleVar(value=0.8)
        if audio.is_audio_initialized():
            audio.set_volume(self.volume_var.get())

        # minimum duration (seconds) configurable via UI
        self.min_duration_var = tk.IntVar(value=MIN_DURATION_SECONDS)
        # string var for entry widget so we can accept free text and validate on submit
        self.min_duration_strvar = tk.StringVar(value=str(self.min_duration_var.get()))
        # dark mode toggle
        self.dark_mode_var = tk.BooleanVar(value=False)

        self.songs_dir = get_default_osu_songs_dir()
        # diagnostic: print songs_dir info
        try:
            print(f"Osu songs dir: {self.songs_dir} (exists={self.songs_dir.exists()})")
            if self.songs_dir.exists():
                try:
                    count = sum(1 for _ in self.songs_dir.iterdir())
                    print(f"  Contains {count} items")
                except Exception:
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

        # Create a container so we can place vertical and horizontal scrollbars correctly
        list_container = ttk.Frame(left)
        list_container.pack(fill=tk.BOTH, expand=True)

        # listbox inside container using grid so hscroll sits under list and vscroll to right
        self.listbox = tk.Listbox(list_container, activestyle='none', exportselection=False)
        self.listbox.grid(row=0, column=0, sticky='nsew')
        self.listbox.bind('<Double-1>', self.on_double_click)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        # Right-click context menu for adding to playlist
        try:
            self.listbox.bind('<Button-3>', self._on_song_right_click)
        except Exception:
            pass

        # vertical scrollbar
        try:
            scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.listbox.yview)
        except Exception:
            scrollbar = tk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.listbox.config(yscrollcommand=scrollbar.set)

        # horizontal scrollbar beneath
        try:
            self.hscroll = ttk.Scrollbar(list_container, orient=tk.HORIZONTAL, command=self.listbox.xview)
        except Exception:
            self.hscroll = tk.Scrollbar(list_container, orient=tk.HORIZONTAL, command=self.listbox.xview)
        self.hscroll.grid(row=1, column=0, columnspan=2, sticky='ew')
        self.listbox.config(xscrollcommand=self.hscroll.set)

        # Make grid expand
        try:
            list_container.rowconfigure(0, weight=1)
            list_container.columnconfigure(0, weight=1)
        except Exception:
            pass

        # Lightweight tooltip for showing full title on hover with delay
        self._title_tooltip = None
        self._tooltip_after_id = None
        self._last_tooltip_index = None
        self._tooltip_delay_ms = 400
        self.listbox.bind('<Motion>', self._on_listbox_motion)
        self.listbox.bind('<Leave>', self._hide_title_tooltip)

        right = ttk.Frame(mid, width=480)
        right.pack(side=tk.RIGHT, fill=tk.Y)
        # keep a reference for placing compact panels like Playlists
        self.right_panel = right
        # Background thumbnail (fixed-size to avoid layout shifts)
        self.meta_image_label = ttk.Label(right)
        self.meta_image_label.pack(anchor=tk.CENTER, padx=6, pady=6)

        # Metadata labels: reserve two lines by default; longer per line; wrap safely
        self._meta_label_width = 104  # approx chars per line (doubled)
        self.meta_title = ttk.Label(right, text="Title: ", width=self._meta_label_width, wraplength=440, justify=tk.LEFT)
        self.meta_title.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_artist = ttk.Label(right, text="Artist: ", width=self._meta_label_width, wraplength=440, justify=tk.LEFT)
        self.meta_artist.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_album = ttk.Label(right, text="Album: ", width=self._meta_label_width, wraplength=440, justify=tk.LEFT)
        self.meta_album.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_duration = ttk.Label(right, text="Duration: ", width=self._meta_label_width, wraplength=440, justify=tk.LEFT)
        self.meta_duration.pack(anchor=tk.W, padx=6, pady=4)
        self.meta_path = ttk.Label(right, text="Path: ", width=self._meta_label_width, wraplength=440, justify=tk.LEFT)
        self.meta_path.pack(anchor=tk.W, padx=6, pady=4)
        try:
            self.meta_path.bind('<Enter>', self._on_meta_path_enter)
            self.meta_path.bind('<Leave>', self._on_meta_path_leave)
        except Exception:
            pass

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
        # Prepare fixed-size placeholder images to prevent layout shifts when images change
        try:
            self._now_img_size = (120, 80)
            self._meta_img_size = (220, 140)
            # Create placeholders using PIL when available, else Tk PhotoImage
            def _mk_placeholder(size):
                w, h = size
                try:
                    return tk.PhotoImage(width=w, height=h)
                except Exception:
                    return None
            self._now_placeholder = _mk_placeholder(self._now_img_size)
            self._meta_placeholder = _mk_placeholder(self._meta_img_size)
            if self._now_placeholder is not None:
                self.now_image_label.config(image=self._now_placeholder)
                setattr(self.now_image_label, '_photo_ref', self._now_placeholder)
            if self._meta_placeholder is not None:
                self.meta_image_label.config(image=self._meta_placeholder)
                setattr(self.meta_image_label, '_photo_ref', self._meta_placeholder)
        except Exception:
            pass
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

        self.skip_btn = ttk.Button(bottom, text="Skip", command=self.skip_track)
        self.skip_btn.pack(side=tk.LEFT)

        # play mode button: 'sequential', 'loop' (repeat current), 'shuffle' (random next)
        self.play_mode = 'sequential'  # default: advance to next
        try:
            self.mode_btn = ttk.Button(bottom, text="Mode: Sequential", command=self.cycle_play_mode)
            self.mode_btn.pack(side=tk.LEFT, padx=(6, 0))
        except Exception:
            try:
                self.mode_btn = tk.Button(bottom, text="Mode: Sequential", command=self.cycle_play_mode)
                self.mode_btn.pack(side=tk.LEFT, padx=(6, 0))
            except Exception:
                self.mode_btn = None

        # Volume control
        self.volume_label = ttk.Label(bottom, text=f"Vol: {int(self.volume_var.get()*100)}%", width=10)
        self.volume_label.pack(side=tk.LEFT, padx=(8, 4))
        # Use a ttk.Scale for volume (0.0 - 1.0)
        self.volume_scale = ttk.Scale(bottom, from_=0.0, to=1.0, orient=tk.HORIZONTAL,
                          length=120, variable=self.volume_var,
                          command=self.on_volume_change)
        self.volume_scale.pack(side=tk.LEFT)

        # Minimum duration entry
        try:
            self.min_label = ttk.Label(bottom, text="Min length (s):")
            self.min_label.pack(side=tk.LEFT, padx=(8, 4))
            self.min_entry = ttk.Entry(bottom, textvariable=self.min_duration_strvar, width=8)
            self.min_entry.pack(side=tk.LEFT)
            # on Enter or focus-out, validate and trigger a background rescan
            self.min_entry.bind('<Return>', lambda e: threading.Thread(target=self._on_min_duration_changed, daemon=True).start())
            self.min_entry.bind('<FocusOut>', lambda e: threading.Thread(target=self._on_min_duration_changed, daemon=True).start())
        except Exception:
            # fallback to simple tk.Entry
            try:
                self.min_label = ttk.Label(bottom, text="Min length (s):")
                self.min_label.pack(side=tk.LEFT, padx=(8, 4))
                self.min_entry = tk.Entry(bottom, textvariable=self.min_duration_strvar, width=8)
                self.min_entry.pack(side=tk.LEFT)
                self.min_entry.bind('<Return>', lambda e: threading.Thread(target=self._on_min_duration_changed, daemon=True).start())
                self.min_entry.bind('<FocusOut>', lambda e: threading.Thread(target=self._on_min_duration_changed, daemon=True).start())
            except Exception:
                pass

        self.current_label = ttk.Label(bottom, text="Not playing")
        self.current_label.pack(side=tk.RIGHT)

        # scan on start (in background)
        self.after(100, lambda: threading.Thread(target=self.scan_and_populate, daemon=True).start())

        self.paused = False
        # playlist playback state flag
        self._playlist_runner_active = False
        self._playlist_cancelled = False
        self._playlist_skip_requested = False
        # metadata cache: path -> dict
        self._metadata = {}
        # counter for excluded short files during scanning (updated on main thread)
        self._excluded_short = 0
        # persistent cache file path
        try:
            self.cache_path = Path.home() / CACHE_FILENAME
        except Exception:
            self.cache_path = Path(CACHE_FILENAME)

        # Playlists store
        try:
            self.playlists = PlaylistStore()
        except Exception:
            self.playlists = PlaylistStore(storage_path=Path(".osu_song_browser_playlists.json"))

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

        # --- Playlists UI ---
        try:
            self._init_playlists_ui()
        except Exception:
            pass

        # Ensure theme is applied after all widgets are created
        try:
            self.after(0, self.apply_theme)
        except Exception:
            pass

        # Build context menu after playlists init
        try:
            self._build_song_context_menu()
        except Exception:
            pass

    def _init_playlists_ui(self):
        """Create a compact playlists section inside the right panel."""
        parent = getattr(self, 'right_panel', self)
        pl_frame = ttk.LabelFrame(parent, text="Playlists")
        pl_frame.pack(fill=tk.X, padx=6, pady=6)

        row = ttk.Frame(pl_frame)
        row.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(row, text="Name:").pack(side=tk.LEFT)
        self.playlist_name_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.playlist_name_var, width=24).pack(side=tk.LEFT, padx=6)
        ttk.Button(row, text="Create", command=self._on_create_playlist).pack(side=tk.LEFT)

        list_row = ttk.Frame(pl_frame)
        list_row.pack(fill=tk.X, padx=6)
        # smaller listbox height to reduce footprint
        self.playlist_listbox = tk.Listbox(list_row, height=4, exportselection=False)
        self.playlist_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        try:
            pl_scroll = ttk.Scrollbar(list_row, orient=tk.VERTICAL, command=self.playlist_listbox.yview)
        except Exception:
            pl_scroll = tk.Scrollbar(list_row, orient=tk.VERTICAL, command=self.playlist_listbox.yview)
        pl_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.playlist_listbox.config(yscrollcommand=pl_scroll.set)
        # refresh tracks when a playlist is selected
        try:
            self.playlist_listbox.bind('<<ListboxSelect>>', self._on_playlist_select)
        except Exception:
            pass

        btns = ttk.Frame(pl_frame)
        btns.pack(fill=tk.X, padx=6, pady=4)
        # Target playlist dropdown to avoid changing selection focus
        ttk.Label(btns, text="Target:").pack(side=tk.LEFT)
        self.playlist_target_var = tk.StringVar()
        self.playlist_target_combo = ttk.Combobox(btns, textvariable=self.playlist_target_var, state='readonly', width=18)
        self.playlist_target_combo.pack(side=tk.LEFT, padx=(4, 8))
        try:
            self.playlist_target_combo.bind('<<ComboboxSelected>>', self._on_target_playlist_changed)
        except Exception:
            pass
        ttk.Button(btns, text="Add Selected Song", command=self._on_add_selected_to_playlist).pack(side=tk.LEFT)
        ttk.Button(btns, text="Play Playlist", command=self._on_play_playlist).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="Delete", command=self._on_delete_playlist).pack(side=tk.RIGHT)

        # Inline status area for playlist actions (avoids popups)
        self._playlist_status_after_id = None
        self.playlist_status = ttk.Label(pl_frame, text="")
        self.playlist_status.pack(fill=tk.X, padx=6, pady=(2, 2))

        # Tracks list for selected playlist
        tracks_lbl = ttk.Label(pl_frame, text="Tracks:")
        tracks_lbl.pack(anchor=tk.W, padx=6, pady=(2, 0))
        tracks_row = ttk.Frame(pl_frame)
        tracks_row.pack(fill=tk.BOTH, expand=False, padx=6)
        self.playlist_tracks_listbox = tk.Listbox(tracks_row, height=6, exportselection=False)
        self.playlist_tracks_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        try:
            tracks_scroll = ttk.Scrollbar(tracks_row, orient=tk.VERTICAL, command=self.playlist_tracks_listbox.yview)
        except Exception:
            tracks_scroll = tk.Scrollbar(tracks_row, orient=tk.VERTICAL, command=self.playlist_tracks_listbox.yview)
        tracks_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.playlist_tracks_listbox.config(yscrollcommand=tracks_scroll.set)
        try:
            self.playlist_tracks_listbox.bind('<Double-1>', self._on_playlist_track_double_click)
        except Exception:
            pass
        try:
            self.playlist_tracks_listbox.bind('<<ListboxSelect>>', self._on_playlist_track_select)
        except Exception:
            pass
        # keep metadata wrap aligned with playlist listbox width
        try:
            self.playlist_tracks_listbox.bind('<Configure>', self._on_playlist_tracks_resize)
        except Exception:
            pass
        # current displayed tracks as paths
        self._current_playlist_tracks = []

        # populate initial list
        self._refresh_playlists_list()

    def _refresh_playlists_list(self):
        try:
            # preserve selection name
            cur_sel_name = self._get_selected_playlist_name()
            self.playlist_listbox.delete(0, tk.END)
            names = self.playlists.list_names() if self.playlists else []
            for n in names:
                self.playlist_listbox.insert(tk.END, n)
            # restore selection if possible
            if cur_sel_name and cur_sel_name in names:
                try:
                    idx = names.index(cur_sel_name)
                    self.playlist_listbox.selection_set(idx)
                    self.playlist_listbox.activate(idx)
                except Exception:
                    pass
            # also refresh dropdown
            try:
                self.playlist_target_combo['values'] = names
                # preserve selection if possible
                cur = self.playlist_target_var.get()
                if cur not in names:
                    self.playlist_target_var.set(names[0] if names else '')
            except Exception:
                pass
            # refresh context menu submenu
            try:
                self._build_song_context_menu()
            except Exception:
                pass
            # refresh tracks view if a playlist is selected
            try:
                cur_name = self._get_selected_playlist_name()
                if cur_name:
                    self._refresh_playlist_tracks(cur_name)
                else:
                    self._refresh_playlist_tracks(None)
            except Exception:
                pass
        except Exception:
            pass

    def _on_create_playlist(self):
        name = (self.playlist_name_var.get() or '').strip()
        if not name:
            self._set_playlist_status("Enter a playlist name")
            return
        try:
            if self.playlists:
                self.playlists.create(name)
                self._refresh_playlists_list()
                self.playlist_name_var.set('')
                self._set_playlist_status(f"Created playlist '{name}'")
        except Exception as e:
            self._set_playlist_status(f"Create failed: {e}")

    def _get_selected_song_path(self):
        try:
            sel = self.listbox.curselection()
            if not sel:
                return None
            idx = sel[0]
            path, _title = self.mp3_paths[idx]
            return str(path)
        except Exception:
            return None

    def _get_selected_playlist_name(self):
        try:
            sel = self.playlist_listbox.curselection()
            if not sel:
                return None
            return self.playlist_listbox.get(sel[0])
        except Exception:
            return None

    def _on_add_selected_to_playlist(self):
        # Prefer dropdown selection; fallback to listbox
        pl_name = (self.playlist_target_var.get() or '').strip()
        if not pl_name:
            pl_name = self._get_selected_playlist_name()
        if not pl_name:
            self._set_playlist_status("Select a playlist (dropdown or list)")
            return
        song = self._get_selected_song_path()
        if not song:
            self._set_playlist_status("Select a song in the list")
            return
        try:
            if self.playlists:
                self.playlists.add_track(pl_name, song)
                self._set_playlist_status(f"Added to '{pl_name}'")
                # refresh track list if editing current playlist
                try:
                    cur_name = self._get_selected_playlist_name()
                    if cur_name == pl_name:
                        self._refresh_playlist_tracks(pl_name)
                except Exception:
                    pass
        except Exception as e:
            self._set_playlist_status(f"Add failed: {e}")

    def _on_delete_playlist(self):
        pl_name = self._get_selected_playlist_name()
        if not pl_name:
            return
        try:
            if self.playlists:
                self.playlists.delete(pl_name)
                self._refresh_playlists_list()
                self._set_playlist_status(f"Deleted '{pl_name}'")
        except Exception as e:
            self._set_playlist_status(f"Delete failed: {e}")

    def _on_play_playlist(self):
        pl_name = self._get_selected_playlist_name()
        if not pl_name:
            self._set_playlist_status("Select a playlist to play")
            return
        pl = self.playlists.get(pl_name) if self.playlists else None
        if not pl or not pl.tracks:
            self._set_playlist_status("Playlist is empty")
            return
        # play sequentially using existing controls
        self._play_playlist_tracks(list(pl.tracks))

    def _on_playlist_select(self, event):
        try:
            name = self._get_selected_playlist_name()
            # keep target dropdown in sync with list selection
            if name:
                try:
                    self.playlist_target_var.set(name)
                except Exception:
                    pass
            self._refresh_playlist_tracks(name)
        except Exception:
            pass

    def _refresh_playlist_tracks(self, name):
        try:
            self.playlist_tracks_listbox.delete(0, tk.END)
            self._current_playlist_tracks = []
            if not name or not self.playlists:
                return
            pl = self.playlists.get(name)
            if not pl:
                return
            for p in pl.tracks:
                try:
                    path = Path(p)
                except Exception:
                    path = None
                display = str(p)
                if path is not None:
                    folder = path.parent.name if path.parent else path.name
                    display = strip_leading_numbers(folder)
                self.playlist_tracks_listbox.insert(tk.END, display)
                self._current_playlist_tracks.append(str(p))
        except Exception:
            pass

    def _on_playlist_track_double_click(self, event):
        try:
            sel = self.playlist_tracks_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            if 0 <= idx < len(self._current_playlist_tracks):
                # Start playlist-mode playback from the double-clicked track and wrap around
                try:
                    # Cancel any existing playlist runner
                    self._playlist_cancelled = True
                except Exception:
                    pass
                try:
                    tracks = list(self._current_playlist_tracks)
                    if not tracks:
                        return
                    self._play_playlist_tracks(tracks, start_index=idx, wrap=True)
                except Exception:
                    # Fallback: play only this track in playlist context
                    p = self._current_playlist_tracks[idx]
                    self._play_path(Path(p), from_playlist=True)
        except Exception:
            pass

    def _on_playlist_track_select(self, event):
        try:
            sel = self.playlist_tracks_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            if 0 <= idx < len(self._current_playlist_tracks):
                p = self._current_playlist_tracks[idx]
                self._update_meta_display(Path(p))
        except Exception:
            pass

    def _select_playlist_track_by_path(self, path: Path):
        """Select the given path in the playlist tracks list if it is displayed."""
        try:
            if not hasattr(self, 'playlist_tracks_listbox') or self.playlist_tracks_listbox is None:
                return
            if not hasattr(self, '_current_playlist_tracks') or not self._current_playlist_tracks:
                return
            target = str(path)
            idx = None
            for i, p in enumerate(self._current_playlist_tracks):
                try:
                    if str(Path(p)) == str(Path(target)):
                        idx = i
                        break
                except Exception:
                    if p == target:
                        idx = i
                        break
            if idx is not None:
                try:
                    self.playlist_tracks_listbox.selection_clear(0, tk.END)
                except Exception:
                    pass
                try:
                    self.playlist_tracks_listbox.selection_set(idx)
                    self.playlist_tracks_listbox.activate(idx)
                    self.playlist_tracks_listbox.see(idx)
                except Exception:
                    pass
        except Exception:
            pass

    def _on_target_playlist_changed(self, event):
        """Sync list selection and tracks view when target combobox changes."""
        try:
            name = (self.playlist_target_var.get() or '').strip()
            if not name:
                return
            # Find the matching index in the playlist listbox and select it
            try:
                items = [self.playlist_listbox.get(i) for i in range(self.playlist_listbox.size())]
                if name in items:
                    idx = items.index(name)
                    self.playlist_listbox.selection_clear(0, tk.END)
                    self.playlist_listbox.selection_set(idx)
                    self.playlist_listbox.activate(idx)
            except Exception:
                pass
            # Refresh tracks for the chosen playlist
            self._refresh_playlist_tracks(name)
        except Exception:
            pass

    def _on_playlist_tracks_resize(self, event=None):
        """Keep metadata wraplength and character width aligned to playlist list box width."""
        try:
            width_px = None
            try:
                if event is not None and hasattr(event, 'width'):
                    width_px = int(event.width)
            except Exception:
                width_px = None
            if width_px is None:
                try:
                    width_px = int(self.playlist_tracks_listbox.winfo_width())
                except Exception:
                    return

            # Update wraplength for all metadata labels
            for lbl in (self.meta_title, self.meta_artist, self.meta_album, self.meta_duration, self.meta_path):
                try:
                    lbl.config(wraplength=width_px)
                except Exception:
                    pass

            # Estimate characters per line from current font
            per_line_chars = None
            try:
                if tkfont is not None:
                    f = tkfont.nametofont('TkDefaultFont')
                    sample = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
                    px = max(1, f.measure(sample))
                    avg = px / len(sample)
                    per_line_chars = max(10, int(width_px / max(1, avg)))
            except Exception:
                per_line_chars = None
            if per_line_chars is None:
                # fallback heuristic: ~8px per char
                per_line_chars = max(10, int(width_px / 8))

            # Apply width in characters for label widgets (stabilize width)
            try:
                self._meta_label_width = per_line_chars
            except Exception:
                pass
            for lbl in (self.meta_title, self.meta_artist, self.meta_album, self.meta_duration, self.meta_path):
                try:
                    lbl.config(width=self._meta_label_width)
                except Exception:
                    pass

            # Refresh current meta texts to honor new widths (if a path is selected/playing)
            try:
                path = getattr(self, '_playing_path', None)
                if path is not None:
                    self._update_meta_display(path)
            except Exception:
                pass
        except Exception:
            pass

    def _set_playlist_status(self, text: str, duration_ms: int = 2500):
        """Show a transient status message in the playlist area without using popups."""
        try:
            if not hasattr(self, 'playlist_status') or self.playlist_status is None:
                return
            self.playlist_status.config(text=text)
            # cancel previous scheduled clear
            try:
                if self._playlist_status_after_id:
                    self.after_cancel(self._playlist_status_after_id)
            except Exception:
                pass
            # schedule clear
            self._playlist_status_after_id = self.after(duration_ms, lambda: self.playlist_status.config(text=""))
        except Exception:
            pass

    def _play_playlist_tracks(self, tracks: list[str], start_index: int | None = None, wrap: bool = True):
        # basic sequential playback using existing audio functions
        def _runner():
            self._playlist_runner_active = True
            self._playlist_cancelled = False
            
            def make_order(first_cycle: bool = False):
                try:
                    base = list(tracks)
                except Exception:
                    base = list(tracks) if tracks else []
                if not base:
                    return []
                # Shuffle mode: if starting from a specific index on first cycle, play that first then shuffle the rest
                if self.play_mode == 'shuffle':
                    if first_cycle and start_index is not None and 0 <= start_index < len(base):
                        try:
                            first = base[start_index]
                            rest = base[:start_index] + base[start_index+1:]
                        except Exception:
                            first = base[0]
                            rest = base[1:]
                        try:
                            random.shuffle(rest)
                        except Exception:
                            pass
                        return [first] + rest
                    # subsequent cycles or no start index: pure shuffle of full list
                    try:
                        random.shuffle(base)
                    except Exception:
                        pass
                    return base
                # Sequential mode
                if first_cycle and start_index is not None and 0 <= start_index < len(base):
                    return base[start_index:] + base[:start_index]
                return base

            # Build initial order (respect start index on first cycle only)
            order = make_order(first_cycle=True)
            last_started = None
            while not self._playlist_cancelled:
                for p in order:
                    if self._playlist_cancelled:
                        break
                    try:
                        # Start playback on the main thread via _play_path to ensure
                        # progress bar and thumbnail update correctly.
                        started = threading.Event()
                        def _start_on_main(path_str=p):
                            try:
                                self._play_path(Path(path_str), from_playlist=True)
                            finally:
                                started.set()
                        self.after(0, _start_on_main)
                        started.wait(timeout=2.0)
                        last_started = Path(p)
                        # Wait until track finishes; do not advance while paused
                        while True:
                            try:
                                if self._playlist_skip_requested:
                                    # consume skip and stop current, advance to next
                                    self._playlist_skip_requested = False
                                    try:
                                        audio.stop()
                                    except Exception:
                                        pass
                                    time.sleep(0.05)
                                    break
                                if self._playlist_cancelled:
                                    break
                                if getattr(self, 'paused', False):
                                    time.sleep(0.1)
                                    continue
                                if not audio.is_busy():
                                    break
                                time.sleep(0.2)
                            except Exception:
                                break
                    except Exception:
                        continue
                if not wrap:
                    break
                # Build next cycle order: ignore start_index from here on
                order = make_order(first_cycle=False)
            # reset when done
            self._playlist_runner_active = False
            if not self._playlist_cancelled:
                # Only clear playing state if we finished naturally
                try:
                    if last_started is None or self._playing_path == last_started:
                        self._playing_path = None
                        self.after(0, lambda: self.current_label.config(text="Not playing"))
                except Exception:
                    pass
        threading.Thread(target=_runner, daemon=True).start()

    def _update_now_labels(self, path: Path):
        try:
            # Use folder name formatting like the main list
            folder = path.parent.name if path.parent else path.name
            display = strip_leading_numbers(folder)
            self.now_title_label.config(text=f"Now: {display}")
            self.current_label.config(text=display)
        except Exception:
            pass

    # --- Meta label formatting helpers to prevent layout shifts ---
    def _ellipsize_end(self, text: str, max_chars: int) -> str:
        try:
            s = str(text or '')
            if max_chars is None or len(s) <= max_chars:
                return s
            if max_chars <= 1:
                return '…'
            return s[: max_chars - 1] + '…'
        except Exception:
            return str(text)

    def _ellipsize_middle(self, text: str, max_chars: int) -> str:
        try:
            s = str(text or '')
            if max_chars is None or len(s) <= max_chars:
                return s
            if max_chars <= 1:
                return '…'
            head = (max_chars - 1) // 2
            tail = (max_chars - 1) - head
            return s[:head] + '…' + s[-tail:]
        except Exception:
            return str(text)

    def _format_meta_line(self, prefix: str, value: str, max_chars: int, middle: bool = False) -> str:
        try:
            body = self._ellipsize_middle(value, max_chars) if middle else self._ellipsize_end(value, max_chars)
            return f"{prefix}{body}"
        except Exception:
            return f"{prefix}{value}"

    def _format_meta_two_lines(self, prefix: str, value: str, per_line_chars: int, middle: bool = False) -> str:
        """Return up to two lines of text with optional ellipsis.
        - per_line_chars: approx characters per line
        - middle=True uses middle ellipsis on second line (useful for paths)
        Always returns exactly two lines (second may be empty) to stabilize layout.
        """
        try:
            s = str(value or '')
            avail1 = max(0, per_line_chars - len(prefix))
            # Make second line the same length as the top line's content area
            avail2 = max(0, per_line_chars - len(prefix))
            if not s:
                return f"{prefix}\n"

            if middle:
                if len(s) <= avail1:
                    return f"{prefix}{s}\n"
                if len(s) <= avail1 + avail2:
                    l1 = s[:avail1]
                    l2 = s[avail1:]
                    return f"{prefix}{l1}\n{l2}"
                l1 = s[:avail1]
                tail_len = max(0, avail2 - 1)
                l2 = ('…' + s[-tail_len:]) if tail_len > 0 else '…'
                return f"{prefix}{l1}\n{l2}"
            else:
                if len(s) <= avail1:
                    return f"{prefix}{s}\n"
                if len(s) <= avail1 + avail2:
                    l1 = s[:avail1]
                    l2 = s[avail1:]
                    return f"{prefix}{l1}\n{l2}"
                l1 = s[:avail1]
                l2 = s[avail1:avail1 + max(0, avail2 - 1)] + '…'
                return f"{prefix}{l1}\n{l2}"
        except Exception:
            try:
                return f"{prefix}{value}\n"
            except Exception:
                return f"{prefix}\n"

    # --- Context menu for songs ---
    def _build_song_context_menu(self):
        try:
            # Base menu
            self.song_menu = tk.Menu(self, tearoff=0)
            # Submenu for playlists
            self.song_menu_playlists = tk.Menu(self.song_menu, tearoff=0)
            names = self.playlists.list_names() if self.playlists else []
            if names:
                for name in names:
                    self.song_menu_playlists.add_command(
                        label=name,
                        command=lambda n=name: self._add_current_hover_to_playlist(n)
                    )
            else:
                self.song_menu_playlists.add_command(label="No playlists", state=tk.DISABLED)
            self.song_menu.add_cascade(label="Add to playlist", menu=self.song_menu_playlists)
        except Exception:
            pass

    def _on_song_right_click(self, event):
        try:
            # Select the row under mouse
            index = self.listbox.nearest(event.y)
            if index is not None:
                try:
                    self.listbox.selection_clear(0, tk.END)
                except Exception:
                    pass
                self.listbox.selection_set(index)
                self.listbox.activate(index)
                # Track hover index for adding via context menu
                self._last_hover_index = index
            # Show the menu
            try:
                self.song_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.song_menu.grab_release()
        except Exception:
            pass

    def _add_current_hover_to_playlist(self, playlist_name: str):
        try:
            idx = getattr(self, '_last_hover_index', None)
            if idx is None:
                # fallback to current selection
                sel = self.listbox.curselection()
                if not sel:
                    return
                idx = sel[0]
            path, _title = self.mp3_paths[idx]
            if self.playlists:
                self.playlists.add_track(playlist_name, str(path))
                messagebox.showinfo("Playlist", f"Added to '{playlist_name}'.")
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
                    searchable = [folder_title.lower(), str(meta.get('title', '')).lower(), str(meta.get('artist', '')).lower()]
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
            # update play mode button label to reflect loaded mode
            try:
                if self.mode_btn and self.play_mode:
                    mode_text = self.play_mode.capitalize()
                    self.mode_btn.config(text=f"Mode: {mode_text}")
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
                border = '#3f3f3f'
            else:
                # explicit light palette (avoid None to prevent type issues)
                bg = '#f0f0f0'
                fg = '#000000'
                entry_bg = '#ffffff'
                list_bg = '#ffffff'
                select_bg = '#3399ff'
                button_bg = '#e0e0e0'
                border = '#c9c9c9'

            # configure ttk styles
            try:
                style.configure('TFrame', background=bg)
                style.configure('TLabel', background=bg, foreground=fg)
                style.configure('TButton', background=button_bg, foreground=fg)
                style.configure('TEntry', fieldbackground=entry_bg, foreground=fg)
                # Combobox styling
                style.configure('TCombobox', fieldbackground=entry_bg, foreground=fg, background=bg)
                style.configure('Playlist.TCombobox', fieldbackground=entry_bg, foreground=fg, background=bg)
                try:
                    style.map('Playlist.TCombobox',
                              fieldbackground=[('readonly', entry_bg), ('!readonly', entry_bg)],
                              foreground=[('disabled', '#888888'), ('!disabled', fg)],
                              background=[('active', button_bg), ('!active', bg)])
                except Exception:
                    pass
                style.configure('TLabelframe', background=bg, foreground=fg, bordercolor=border)
                style.configure('TLabelframe.Label', background=bg, foreground=fg)
                style.configure('Horizontal.TScale', background=bg)
                style.configure('TScrollbar', background=bg)
                # progressbar styling (may vary by platform)
                try:
                    style.configure('Horizontal.TProgressbar', background='#4CAF50')
                except Exception:
                    pass
            except Exception:
                pass

            # Ensure ttk Combobox dropdown list matches theme via option database
            try:
                self.option_add('*TCombobox*Listbox.background', list_bg)
                self.option_add('*TCombobox*Listbox.foreground', fg)
                self.option_add('*TCombobox*Listbox.selectBackground', select_bg)
                self.option_add('*TCombobox*Listbox.selectForeground', fg)
                self.option_add('*TCombobox*Listbox.highlightBackground', bg)
            except Exception:
                pass

            # apply to some direct tk widgets
            try:
                self.configure(bg=bg)
            except Exception:
                pass
            # Apply custom style to target combobox, if present
            try:
                if hasattr(self, 'playlist_target_combo') and self.playlist_target_combo:
                    self.playlist_target_combo.configure(style='Playlist.TCombobox')
            except Exception:
                pass
            try:
                # Ensure listbox interior matches theme fully (set options individually)
                self.listbox.config(bg=list_bg)
                self.listbox.config(fg=fg)
                self.listbox.config(selectbackground=select_bg)
                self.listbox.config(selectforeground=fg)
                self.listbox.config(highlightbackground=bg)
                self.listbox.config(highlightcolor=bg)
                # skip insertbackground to avoid type-check complaints
            except Exception:
                pass
            try:
                if hasattr(self, 'playlist_listbox') and self.playlist_listbox:
                    self.playlist_listbox.config(bg=list_bg)
                    self.playlist_listbox.config(fg=fg)
                    self.playlist_listbox.config(selectbackground=select_bg)
                    self.playlist_listbox.config(selectforeground=fg)
                    self.playlist_listbox.config(highlightbackground=bg)
                    self.playlist_listbox.config(highlightcolor=bg)
                    # skip insertbackground to avoid type-check complaints
            except Exception:
                pass
            try:
                if hasattr(self, 'playlist_tracks_listbox') and self.playlist_tracks_listbox:
                    self.playlist_tracks_listbox.config(bg=list_bg)
                    self.playlist_tracks_listbox.config(fg=fg)
                    self.playlist_tracks_listbox.config(selectbackground=select_bg)
                    self.playlist_tracks_listbox.config(selectforeground=fg)
                    self.playlist_tracks_listbox.config(highlightbackground=bg)
                    self.playlist_tracks_listbox.config(highlightcolor=bg)
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
                if settings.get('play_mode'):
                    self.play_mode = settings.get('play_mode')
            except Exception:
                pass

            # validate and load
            self.all_mp3_paths.clear()
            for rec in items:
                try:
                    p = Path(rec['path'])
                    if not p.exists():
                        continue
                    folder_title = rec.get('folder_title') or strip_leading_numbers(p.parent.name)
                    self.all_mp3_paths.append((p, folder_title))
                    # restore metadata
                    meta = rec.get('meta', {})
                    if meta:
                        self._metadata[str(p)] = meta
                except Exception:
                    pass
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
                    rec = {'path': str(p), 'folder_title': folder_title}
                    meta = self._metadata.get(str(p))
                    if meta:
                        rec['meta'] = meta
                    out.append(rec)
                except Exception:
                    pass
            try:
                settings = {}
                try:
                    settings['dark_mode'] = bool(self.dark_mode_var.get())
                except Exception:
                    settings['dark_mode'] = False
                try:
                    settings['play_mode'] = self.play_mode
                except Exception:
                    settings['play_mode'] = 'sequential'
                payload = {'items': out, 'settings': settings}
                with self.cache_path.open('w', encoding='utf-8') as f:
                    json.dump(payload, f)
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
        """Called when the min duration entry changes: trigger a re-scan so UI reflects new cutoff."""
        try:
            # parse user input from string var, update IntVar with a safe integer
            try:
                s = (self.min_duration_strvar.get() or '').strip()
                if s == '':
                    val = MIN_DURATION_SECONDS
                else:
                    val = int(s)
            except Exception:
                val = MIN_DURATION_SECONDS
            try:
                self.min_duration_var.set(val)
                # keep the string in sync (normalize formatting)
                self.min_duration_strvar.set(str(val))
            except Exception:
                pass
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
                    dur = ensure_duration(full, self._metadata)
                except Exception:
                    dur = 0
            min_d = self.min_duration_var.get() if hasattr(self, 'min_duration_var') else MIN_DURATION_SECONDS
            if dur and dur < min_d:
                # count as excluded and do not add
                try:
                    self.after(0, self._inc_excluded_short)
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
                    self._metadata[key] = {**self._metadata.get(key, {}), **meta}
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
                    if meta and meta.get('title'):
                        searchable.append(str(meta.get('title')).lower())
                except Exception:
                    pass
                try:
                    if meta and meta.get('artist'):
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
                    self.current_label.config(text=f"Found {len(self.all_mp3_paths)} audio files (excluded {self._excluded_short} < {min_d}s)")
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
                    if any(fn.lower().endswith(ext) for ext in SUPPORTED_AUDIO_EXTS):
                        first_fn = fn
                        break
                if not first_fn:
                    continue
                full = Path(root) / first_fn
                # try to reuse cached metadata if file unchanged
                key = str(full)
                meta = {}
                try:
                    mtime = full.stat().st_mtime
                    size = full.stat().st_size
                except Exception:
                    mtime = None
                    size = None

                cached = self._metadata.get(key)
                if cached and mtime is not None and size is not None and cached.get('__mtime') == mtime and cached.get('__size') == size:
                    meta = cached
                    # already have duration and tags from cache
                else:
                    meta = get_mp3_metadata(full)
                    if mtime is not None:
                        meta['__mtime'] = mtime
                    if size is not None:
                        meta['__size'] = size
                # ensure duration is known (may compute and cache). Prefer cached value.
                dur = meta.get('duration') or 0
                if not dur:
                    try:
                        dur = ensure_duration(full, self._metadata)
                        if dur:
                            meta['duration'] = dur
                    except Exception:
                        dur = 0
                # skip very short files (count as excluded)
                min_d = self.min_duration_var.get() if hasattr(self, 'min_duration_var') else MIN_DURATION_SECONDS
                if dur and dur < min_d:
                    excluded_short += 1
                    continue

                folder_title = strip_leading_numbers(full.parent.name)
                local_all.append((full, folder_title))
                # add this file to the UI immediately
                try:
                    self.after(0, lambda p=full, t=folder_title, m=meta: self._add_discovered_file(p, t, m))
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
                        self.current_label.config(text="No audio files found")
                    except Exception:
                        pass
                if excluded_short:
                    try:
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
                    self._save_cache()
                except Exception:
                    pass
            except Exception as e:
                print(f"Error applying scan results: {e}")

        self.after(0, apply_results)

    def play_selected(self):
        # Prefer main list selection; fallback to playlist track selection
        idx = self.listbox.curselection()
        if idx:
            index = idx[0]
            try:
                path = self.mp3_paths[index][0]
            except IndexError:
                return
            self._play_path(path)
            return
        try:
            tr_sel = self.playlist_tracks_listbox.curselection()
        except Exception:
            tr_sel = ()
        if tr_sel:
            tindex = tr_sel[0]
            if 0 <= tindex < len(self._current_playlist_tracks):
                self._play_path(Path(self._current_playlist_tracks[tindex]))
                return
        messagebox.showinfo("Select", "Please select an audio file from the list or a playlist track.")

    def on_double_click(self, event):
        # On double-click, always play from the main list selection
        idx = self.listbox.curselection()
        if not idx:
            return
        index = idx[0]
        try:
            path = self.mp3_paths[index][0]
        except IndexError:
            return
        self._play_path(path)

    def _play_path(self, path: Path, from_playlist: bool = False):
        try:
            # Cancel any ongoing playlist runner when this is a manual play
            if not from_playlist:
                self._playlist_cancelled = True
                self._playlist_runner_active = False
            if not audio.load_and_play(str(path)):
                messagebox.showerror("Playback error", f"Failed to play {path}")
                return
            
            # display folder title as the song name
            folder_title = strip_leading_numbers(path.parent.name)
            # if we stored folder title in mp3_paths, prefer that
            for p, t in self.mp3_paths:
                if p == path:
                    folder_title = t
                    break
            self.current_label.config(text=f"Playing: {folder_title}")
            # Update now-playing display (thumbnail + title)
            self.now_title_label.config(text=f"Now: {folder_title}")
            # also update the right-side metadata panel to reflect the playing file
            try:
                self._update_meta_display(path)
            except Exception:
                pass
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
            if bg and HAS_PIL and Image and ImageTk:
                try:
                    img = Image.open(bg)
                    # fit into fixed canvas and center to avoid layout shifts
                    canvas_w, canvas_h = self._now_img_size if hasattr(self, '_now_img_size') else (120, 80)
                    resampling = getattr(Image, 'Resampling', None)
                    resample = getattr(resampling, 'LANCZOS', None) if resampling is not None else getattr(Image, 'LANCZOS', None)
                    if resample is not None:
                        img.thumbnail((canvas_w, canvas_h), resample)
                    else:
                        img.thumbnail((canvas_w, canvas_h))
                    # paste onto fixed-size background
                    base = Image.new('RGB', (canvas_w, canvas_h))
                    try:
                        x = (canvas_w - img.width) // 2
                        y = (canvas_h - img.height) // 2
                        base.paste(img, (x, y))
                    except Exception:
                        base = img
                    photo = ImageTk.PhotoImage(base)
                    self.now_image_label.config(image=photo)
                    setattr(self.now_image_label, '_photo_ref', photo)
                except Exception:
                    # fallback to placeholder
                    placeholder = getattr(self, '_now_placeholder', None)
                    if placeholder is not None:
                        self.now_image_label.config(image=placeholder)
                        setattr(self.now_image_label, '_photo_ref', placeholder)
            else:
                # no image available; show fixed-size placeholder
                placeholder = getattr(self, '_now_placeholder', None)
                if placeholder is not None:
                    self.now_image_label.config(image=placeholder)
                    setattr(self.now_image_label, '_photo_ref', placeholder)
            self.paused = False
        except Exception as e:
            messagebox.showerror("Playback error", f"Failed to play {path}: {e}")
        finally:
            # Reflect current playing track selection in playlist tracks list if present
            try:
                self._select_playlist_track_by_path(path)
            except Exception:
                pass

    def toggle_pause(self):
        if not audio.is_audio_initialized():
            return
        # If nothing is playing, do nothing
        if not self._playing_path:
            return

        if not self.paused:
            audio.pause()
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
            unpaused = audio.unpause()
            if not unpaused:
                # compute paused position from manual timer
                pos_sec = 0
                try:
                    if self._start_time and self._pause_time:
                        pos_sec = (self._pause_time - self._start_time) + self._paused_offset
                except Exception:
                    pos_sec = 0
                try:
                    audio.restart_playback(str(self._playing_path))
                    self.seek_to(pos_sec)
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
                    pause_duration = time.time() - self._pause_time
                    self._start_time += pause_duration
            except Exception:
                pass
            self._pause_time = None

    def skip_track(self):
        """Skip to the next track based on current play mode."""
        try:
            # If a playlist runner is active, request skip within playlist
            if getattr(self, '_playlist_runner_active', False):
                self._playlist_skip_requested = True
                try:
                    audio.stop()
                except Exception:
                    pass
                # progress updater will be restarted by next track via _play_path
                if self._progress_after_id:
                    try:
                        self.after_cancel(self._progress_after_id)
                    except Exception:
                        pass
                    self._progress_after_id = None
                return

            # Otherwise, behave as before: cancel updater and go to next in visible list
            if self._progress_after_id:
                try:
                    self.after_cancel(self._progress_after_id)
                except Exception:
                    pass
                self._progress_after_id = None
            self._on_track_end(force_next=True)
        except Exception:
            pass

    def stop(self):
        audio.stop()
        self.current_label.config(text="Not playing")
        # clear now-playing and cancel progress updates
        self._playing_path = None
        self._playlist_runner_active = False
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
        # restore placeholder to keep layout stable
        placeholder = getattr(self, '_now_placeholder', None)
        if placeholder is not None:
            self.now_image_label.config(image=placeholder)
            setattr(self.now_image_label, '_photo_ref', placeholder)
        self.progress['value'] = 0
        self.time_label.config(text="0:00 / 0:00")

    def toggle_loop(self):
        """Toggle looping of the current song. When enabled, the current track will replay after ending."""
        try:
            # kept for backwards-compat; map into play_mode
            if self.play_mode == 'loop':
                self.play_mode = 'sequential'
            else:
                self.play_mode = 'loop'
            try:
                if self.mode_btn:
                    mode_text = self.play_mode.capitalize()
                    self.mode_btn.config(text=f"Mode: {mode_text}")
            except Exception:
                pass
        except Exception:
            pass

    def cycle_play_mode(self):
        """Cycle play mode between 'sequential' -> 'loop' -> 'shuffle' -> sequential."""
        try:
            if self.play_mode == 'sequential':
                self.play_mode = 'loop'
            elif self.play_mode == 'loop':
                self.play_mode = 'shuffle'
            else:
                self.play_mode = 'sequential'
            # update button text
            try:
                if self.mode_btn:
                    mode_text = self.play_mode.capitalize()
                    self.mode_btn.config(text=f"Mode: {mode_text}")
            except Exception:
                pass
            # persist mode into cache
            try:
                self._save_cache()
            except Exception:
                pass
        except Exception:
            pass

    def _on_track_end(self, force_next: bool = False):
        """Called when the current track finishes playing. Decide whether to loop or play next.
        If force_next is True (e.g., Skip pressed), bypass loop mode and advance.
        """
        try:
            # if loop enabled, restart same track
            if self.play_mode == 'loop' and self._playing_path and not force_next:
                try:
                    audio.restart_playback(str(self._playing_path))
                except Exception:
                    pass
                try:
                    self._start_time = time.time()
                except Exception:
                    self._start_time = None
                self._pause_time = None
                self._paused_offset = 0.0
                # continue progress polling
                try:
                    if self._progress_after_id:
                        try:
                            self.after_cancel(self._progress_after_id)
                        except Exception:
                            pass
                    self._progress_after_id = self.after(500, self.update_progress)
                except Exception:
                    pass
                return

            # otherwise, play the next visible song (in self.mp3_paths)
            try:
                if not self._playing_path:
                    return

                # If shuffle mode, pick a random song from the whole library (`all_mp3_paths`).
                if self.play_mode == 'shuffle':
                    if not self.all_mp3_paths:
                        return
                    # avoid picking the same track if possible
                    candidates = [p for p, t in self.all_mp3_paths if p != self._playing_path]
                    if not candidates:
                        candidates = [p for p, t in self.all_mp3_paths]
                    if candidates:
                        next_path = random.choice(candidates)
                        try:
                            # find the matching entry in mp3_paths to highlight in UI
                            for i, (p, t) in enumerate(self.mp3_paths):
                                if p == next_path:
                                    try:
                                        self.listbox.selection_clear(0, tk.END)
                                        self.listbox.selection_set(i)
                                        self.listbox.see(i)
                                    except Exception:
                                        pass
                                    break
                            self._play_path(next_path)
                        except Exception:
                            pass
                    return

                # otherwise behave sequentially within visible list
                # find current index in visible list (mp3_paths)
                cur_index = None
                for i, (p, t) in enumerate(self.mp3_paths):
                    if p == self._playing_path:
                        cur_index = i
                        break
                if cur_index is None:
                    # playing file not in visible list, stop
                    self.stop()
                    return
                next_index = cur_index + 1
                if next_index >= len(self.mp3_paths):
                    # end of list
                    self.stop()
                    return
                # select and play next
                next_path = self.mp3_paths[next_index][0]
                try:
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(next_index)
                    self.listbox.see(next_index)
                    self._play_path(next_path)
                except Exception:
                    pass
            except Exception:
                pass
        except Exception:
            pass

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
        audio.set_volume(v)
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
        if not artist:
            # parse artist from folder name if not present in tags
            artist = parse_artist_from_folder(title) or ''
            # persist parsed artist into metadata cache so it is available later
            try:
                key = str(path)
                meta_entry = self._metadata.get(key, {})
                if not meta_entry.get('artist'):
                    meta_entry['artist'] = artist
                    self._metadata[key] = meta_entry
            except Exception:
                pass
        album = meta.get('album') or ''
        duration = format_duration(meta.get('duration')) if meta.get('duration') else ''
        per_line = getattr(self, '_meta_label_width', 52)
        self.meta_title.config(text=self._format_meta_two_lines('Title: ', title, per_line))
        self.meta_artist.config(text=self._format_meta_two_lines('Artist: ', artist, per_line))
        self.meta_album.config(text=self._format_meta_two_lines('Album: ', album, per_line))
        # Duration: keep two-line shape for stability
        self.meta_duration.config(text=self._format_meta_two_lines('Duration: ', duration, per_line))
        self.meta_path.config(text=self._format_meta_two_lines('Path: ', str(path), per_line, middle=True))
        # Store the full path for tooltip
        try:
            self._meta_path_full = str(path)
        except Exception:
            self._meta_path_full = ''
        # Try to load background from the first .osu file in the folder
        bg = get_osu_background(path.parent)
        if bg and HAS_PIL and Image and ImageTk:
            try:
                img = Image.open(bg)
                # fit into fixed canvas and center (220x140) to prevent layout shifts
                canvas_w, canvas_h = self._meta_img_size if hasattr(self, '_meta_img_size') else (220, 140)
                resampling = getattr(Image, 'Resampling', None)
                resample = getattr(resampling, 'LANCZOS', None) if resampling is not None else getattr(Image, 'LANCZOS', None)
                if resample is not None:
                    img.thumbnail((canvas_w, canvas_h), resample)
                else:
                    img.thumbnail((canvas_w, canvas_h))
                base = Image.new('RGB', (canvas_w, canvas_h))
                try:
                    x = (canvas_w - img.width) // 2
                    y = (canvas_h - img.height) // 2
                    base.paste(img, (x, y))
                except Exception:
                    base = img
                photo = ImageTk.PhotoImage(base)
                self.meta_image_label.config(image=photo)
                setattr(self.meta_image_label, '_photo_ref', photo)
            except Exception:
                # fallback to placeholder
                placeholder = getattr(self, '_meta_placeholder', None)
                if placeholder is not None:
                    self.meta_image_label.config(image=placeholder)
                    setattr(self.meta_image_label, '_photo_ref', placeholder)
        else:
            # show placeholder if none found or PIL missing
            placeholder = getattr(self, '_meta_placeholder', None)
            if placeholder is not None:
                self.meta_image_label.config(image=placeholder)
                setattr(self.meta_image_label, '_photo_ref', placeholder)

    def _update_meta_display(self, path: Path):
        """Update the right-side metadata panel (title/artist/album/duration/path/image) for `path`."""
        try:
            meta = self._metadata.get(str(path), {})
            # display song name based on folder name
            title = strip_leading_numbers(path.parent.name)
            artist = meta.get('artist') or ''
            if not artist:
                artist = parse_artist_from_folder(title) or ''
                # persist parsed artist into metadata cache
                try:
                    key = str(path)
                    meta_entry = self._metadata.get(key, {})
                    if not meta_entry.get('artist'):
                        meta_entry['artist'] = artist
                        self._metadata[key] = meta_entry
                except Exception:
                    pass
            album = meta.get('album') or ''
            duration = format_duration(meta.get('duration')) if meta.get('duration') else ''
            try:
                per_line = getattr(self, '_meta_label_width', 52)
                self.meta_title.config(text=self._format_meta_two_lines('Title: ', title, per_line))
                self.meta_artist.config(text=self._format_meta_two_lines('Artist: ', artist, per_line))
                self.meta_album.config(text=self._format_meta_two_lines('Album: ', album, per_line))
                self.meta_duration.config(text=self._format_meta_two_lines('Duration: ', duration, per_line))
                self.meta_path.config(text=self._format_meta_two_lines('Path: ', str(path), per_line, middle=True))
                # Store full path for tooltip
                try:
                    self._meta_path_full = str(path)
                except Exception:
                    self._meta_path_full = ''
            except Exception:
                pass

            # load background image for meta panel
            bg = get_osu_background(path.parent)
            if bg and HAS_PIL and Image and ImageTk:
                try:
                    img = Image.open(bg)
                    canvas_w, canvas_h = self._meta_img_size if hasattr(self, '_meta_img_size') else (220, 140)
                    resampling = getattr(Image, 'Resampling', None)
                    resample = getattr(resampling, 'LANCZOS', None) if resampling is not None else getattr(Image, 'LANCZOS', None)
                    if resample is not None:
                        img.thumbnail((canvas_w, canvas_h), resample)
                    else:
                        img.thumbnail((canvas_w, canvas_h))
                    base = Image.new('RGB', (canvas_w, canvas_h))
                    try:
                        x = (canvas_w - img.width) // 2
                        y = (canvas_h - img.height) // 2
                        base.paste(img, (x, y))
                    except Exception:
                        base = img
                    photo = ImageTk.PhotoImage(base)
                    self.meta_image_label.config(image=photo)
                    setattr(self.meta_image_label, '_photo_ref', photo)
                except Exception:
                    placeholder = getattr(self, '_meta_placeholder', None)
                    if placeholder is not None:
                        self.meta_image_label.config(image=placeholder)
                        setattr(self.meta_image_label, '_photo_ref', placeholder)
            else:
                try:
                    placeholder = getattr(self, '_meta_placeholder', None)
                    if placeholder is not None:
                        self.meta_image_label.config(image=placeholder)
                        setattr(self.meta_image_label, '_photo_ref', placeholder)
                except Exception:
                    pass
        except Exception:
            pass

    def _on_listbox_motion(self, event):
        """Schedule showing a tooltip near the mouse with the full list item text after a short delay."""
        try:
            lb = event.widget
            idx = lb.nearest(event.y)
            if idx is None:
                self._hide_title_tooltip()
                return
            try:
                text = lb.get(idx)
            except Exception:
                text = ''
            if not text:
                self._hide_title_tooltip()
                return

            # if mouse is still over same index, don't reschedule
            if self._last_tooltip_index == idx and self._title_tooltip:
                # update position if visible
                try:
                    if self._title_tooltip.winfo_exists():
                        x = event.x_root + 12
                        y = event.y_root + 18
                        self._title_tooltip.wm_geometry(f"+{x}+{y}")
                except Exception:
                    pass
                return

            self._last_tooltip_index = idx
            # cancel previous scheduled show
            try:
                if self._tooltip_after_id:
                    self.after_cancel(self._tooltip_after_id)
            except Exception:
                pass

            # schedule showing tooltip after delay
            try:
                x = event.x_root + 12
                y = event.y_root + 18
                self._tooltip_after_id = self.after(self._tooltip_delay_ms, lambda: self._show_title_tooltip(x, y, text, idx))
            except Exception:
                pass
        except Exception:
            pass

    def _hide_title_tooltip(self, event=None):
        try:
            if self._title_tooltip:
                try:
                    self._title_tooltip.destroy()
                except Exception:
                    pass
                self._title_tooltip = None
            # cancel any scheduled show
            try:
                aid = getattr(self, '_tooltip_after_id', None)
                if aid is not None:
                    if hasattr(self, 'after_cancel') and callable(self.after_cancel):
                        self.after_cancel(aid)
            except Exception:
                pass
            self._tooltip_after_id = None
            self._last_tooltip_index = None
        except Exception:
            pass

    def _on_meta_path_enter(self, event=None):
        """Show full path in a tooltip when hovering the Path label."""
        try:
            full = getattr(self, '_meta_path_full', '')
            if not full:
                return
            widget = self.meta_path if hasattr(self, 'meta_path') and self.meta_path is not None else (event.widget if event is not None else None)
            if widget is None:
                return
            x = widget.winfo_rootx() + 10
            y = widget.winfo_rooty() + widget.winfo_height() + 6
            self._show_title_tooltip(x, y, str(full), idx='meta_path')
        except Exception:
            pass

    def _on_meta_path_leave(self, event=None):
        try:
            self._hide_title_tooltip()
        except Exception:
            pass

    def _show_title_tooltip(self, x, y, text, idx):
        """Create and show the tooltip immediately at x,y with given text."""
        try:
            # clear any previous tooltip
            try:
                if self._title_tooltip:
                    self._title_tooltip.destroy()
            except Exception:
                pass

            dark = bool(self.dark_mode_var.get()) if hasattr(self, 'dark_mode_var') else False
            if dark:
                bg = '#222222'
                fg = '#f0f0f0'
            else:
                bg = '#ffffe0'
                fg = '#000000'

            tw = tk.Toplevel(self)
            tw.wm_overrideredirect(True)
            # use tk.Label for easier bg/fg control
            lbl = tk.Label(tw, text=text, bg=bg, fg=fg, bd=1, relief='solid')
            lbl.pack(ipadx=6, ipady=3)
            try:
                tw.wm_geometry(f"+{x}+{y}")
            except Exception:
                pass
            self._title_tooltip = tw
            # clear scheduled id
            self._tooltip_after_id = None
        except Exception:
            pass

    def _clear_search(self):
        self.search_var.set('')
        self.refresh_list()

    def update_progress(self):
        """Poll playback position and update the progress bar and time label."""
        try:
            path = self._playing_path
            if not path or not audio.is_audio_initialized():
                return

            total = self._metadata.get(str(path), {}).get('duration') or 0
            # if duration unknown, try to compute and cache it
            if not total:
                total = ensure_duration(path, self._metadata)

            # Prefer manual timing base for progress display
            busy = audio.is_busy()
            if not busy and not self.paused:
                # If a playlist runner is active, let it handle advancing
                if getattr(self, '_playlist_runner_active', False):
                    try:
                        if self._progress_after_id:
                            self.after_cancel(self._progress_after_id)
                        self._progress_after_id = None
                    except Exception:
                        pass
                    return
                # playback finished; handle end-of-track behavior (loop or advance)
                try:
                    if self._progress_after_id:
                        self.after_cancel(self._progress_after_id)
                    self._progress_after_id = None
                except Exception:
                    pass
                try:
                    self.after(100, self._on_track_end)
                except Exception:
                    self.stop()
                return

            # Compute position using manual base when possible
            pos_sec = 0
            try:
                if self._start_time is not None:
                    if self.paused and self._pause_time:
                        pos_sec = (self._pause_time - self._start_time) + self._paused_offset
                    else:
                        pos_sec = (time.time() - self._start_time) + self._paused_offset
                else:
                    # fallback to pygame get_pos
                    pos_ms = audio.get_pos()
                    pos_sec = pos_ms / 1000.0
            except Exception:
                # fallback to pygame get_pos
                try:
                    pos_ms = audio.get_pos()
                    pos_sec = pos_ms / 1000.0
                except Exception:
                    pos_sec = 0

            if total:
                frac = min(1.0, pos_sec / total)
                self.progress['value'] = int(frac * 1000)
                self.time_label.config(text=f"{format_duration(int(pos_sec))} / {format_duration(total)}")
            else:
                # unknown total
                self.progress['value'] = 0
                self.time_label.config(text=f"{format_duration(int(pos_sec))} / 0:00")

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
            total = self._metadata.get(str(self._playing_path), {}).get('duration') or ensure_duration(self._playing_path, self._metadata)
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
        total = self._metadata.get(str(self._playing_path), {}).get('duration') or ensure_duration(self._playing_path, self._metadata)
        if total and pos_sec > total:
            pos_sec = total
        try:
            # Attempt several seek methods in order for best compatibility
            success = audio.seek_set_pos(pos_sec)
            if not success:
                success = audio.seek_play_start(pos_sec)
            if not success:
                audio.restart_playback(str(self._playing_path))
            
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
