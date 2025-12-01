"""Playlist management for osu! Song Browser.

Provides a simple JSON-backed store of playlists. Each playlist
contains an ordered list of file paths to audio tracks. Paths are
stored as strings. Persistence lives under the user's home directory
alongside the metadata cache.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Dict, Optional


DEFAULT_PLAYLISTS_FILENAME = ".osu_song_browser_playlists.json"


@dataclass
class Playlist:
    name: str
    tracks: List[str] = field(default_factory=list)

    def add(self, path: Path | str) -> None:
        p = str(path)
        if p not in self.tracks:
            self.tracks.append(p)

    def remove(self, path: Path | str) -> None:
        p = str(path)
        try:
            self.tracks.remove(p)
        except ValueError:
            pass

    def clear(self) -> None:
        self.tracks.clear()


class PlaylistStore:
    """Manages loading/saving playlists to a JSON file."""

    def __init__(self, storage_path: Optional[Path] = None):
        if storage_path is None:
            try:
                storage_path = Path.home() / DEFAULT_PLAYLISTS_FILENAME
            except Exception:
                storage_path = Path(DEFAULT_PLAYLISTS_FILENAME)
        self.storage_path = storage_path
        self._playlists: Dict[str, Playlist] = {}
        self.load()

    def load(self) -> None:
        try:
            if self.storage_path.exists():
                data = json.loads(self.storage_path.read_text(encoding="utf-8"))
                self._playlists = {
                    name: Playlist(name=name, tracks=list(pl.get("tracks", [])))
                    for name, pl in dict(data or {}).items()
                }
        except Exception:
            # fail silently and start fresh
            self._playlists = {}

    def save(self) -> None:
        try:
            payload = {name: {"tracks": pl.tracks} for name, pl in self._playlists.items()}
            self.storage_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception:
            pass

    def list_names(self) -> List[str]:
        return sorted(self._playlists.keys())

    def get(self, name: str) -> Optional[Playlist]:
        return self._playlists.get(name)

    def create(self, name: str) -> Playlist:
        name = name.strip()
        if not name:
            raise ValueError("Playlist name cannot be empty")
        if name in self._playlists:
            return self._playlists[name]
        pl = Playlist(name=name)
        self._playlists[name] = pl
        self.save()
        return pl

    def delete(self, name: str) -> None:
        if name in self._playlists:
            del self._playlists[name]
            self.save()

    def add_track(self, name: str, path: Path | str) -> None:
        pl = self.create(name) if name not in self._playlists else self._playlists[name]
        pl.add(path)
        self.save()

    def remove_track(self, name: str, path: Path | str) -> None:
        pl = self._playlists.get(name)
        if not pl:
            return
        pl.remove(path)
        self.save()
