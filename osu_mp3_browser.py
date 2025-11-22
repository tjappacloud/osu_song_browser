import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
import pygame
import sys
import re

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

        self.songs_dir = get_default_osu_songs_dir()
        # store tuples of (Path, folder_title) where folder_title is the parent folder name
        self.all_mp3_paths = []
        self.mp3_paths = []  # list of (Path, display_title)

        # UI
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=6)

        self.dir_label = ttk.Label(top, text=f"Songs dir: {self.songs_dir}")
        self.dir_label.pack(side=tk.LEFT, expand=True)

        browse_btn = ttk.Button(top, text="Browse...", command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT)
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
        self.now_title_label = ttk.Label(now_frame, text="Now: Not playing")
        self.now_title_label.pack(side=tk.LEFT, anchor=tk.CENTER)

        self.play_btn = ttk.Button(bottom, text="Play", command=self.play_selected)
        self.play_btn.pack(side=tk.LEFT)

        self.pause_btn = ttk.Button(bottom, text="Pause", command=self.toggle_pause)
        self.pause_btn.pack(side=tk.LEFT, padx=6)

        self.stop_btn = ttk.Button(bottom, text="Stop", command=self.stop)
        self.stop_btn.pack(side=tk.LEFT)

        self.current_label = ttk.Label(bottom, text="Not playing")
        self.current_label.pack(side=tk.RIGHT)

        # scan on start (in background)
        self.after(100, lambda: threading.Thread(target=self.scan_and_populate, daemon=True).start())

        self.paused = False
        # metadata cache: path -> dict
        self._metadata = {}

    def browse_folder(self):
        path = filedialog.askdirectory(initialdir=str(self.songs_dir) if self.songs_dir.exists() else None)
        if path:
            self.songs_dir = Path(path)
            self.dir_label.config(text=f"Songs dir: {self.songs_dir}")
            threading.Thread(target=self.scan_and_populate, daemon=True).start()

    def scan_and_populate(self):
        self.listbox.delete(0, tk.END)
        self.mp3_paths.clear()
        self.all_mp3_paths.clear()
        if not self.songs_dir.exists():
            self.listbox.insert(tk.END, "(Songs directory not found)")
            return

        # Walk subdirectories and collect mp3 files
        for root, dirs, files in sorted(os_walk(self.songs_dir)):
            for fn in files:
                if fn.lower().endswith('.mp3'):
                    full = Path(root) / fn
                    # read metadata and cache (still useful for details panel)
                    meta = get_mp3_metadata(full) if HAS_MUTAGEN else {}
                    self._metadata[str(full)] = meta
                    # derive display title from parent folder name (strip leading numbers)
                    folder_title = strip_leading_numbers(full.parent.name)
                    self.all_mp3_paths.append((full, folder_title))
                    # visible list will be populated by refresh_list

        # populate visible list from full list (no filter)
        self.refresh_list()
        if not self.all_mp3_paths:
            self.listbox.insert(tk.END, "(No mp3 files found in Songs directory)")

    def play_selected(self):
        idx = self.listbox.curselection()
        if not idx:
            messagebox.showinfo("Select", "Please select an mp3 from the list.")
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
        if pygame.mixer.music.get_busy():
            if not self.paused:
                pygame.mixer.music.pause()
                self.paused = True
                self.pause_btn.config(text="Resume")
                self.current_label.config(text=self.current_label.cget("text") + " (paused)")
            else:
                pygame.mixer.music.unpause()
                self.paused = False
                self.pause_btn.config(text="Pause")
                # remove (paused) suffix
                txt = self.current_label.cget("text").replace(" (paused)", "")
                self.current_label.config(text=txt)

    def stop(self):
        if pygame.mixer.get_init():
            pygame.mixer.music.stop()
        self.current_label.config(text="Not playing")

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
