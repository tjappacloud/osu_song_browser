# osu! Song Browser

A desktop application for browsing and playing audio files from your osu! Songs directory with a feature-rich GUI.

## Features

- **Automatic Song Discovery**: Scans your osu! Songs directory and lists all audio files
- **Smart Metadata Display**: Extracts song information from MP3 tags or folder names
- **Background Thumbnails**: Displays osu! beatmap background images
- **Playback Controls**: Play, pause, stop, seek, and volume control
- **Play Modes**: Sequential, loop current track, or shuffle across entire library
- **Search & Filter**: Real-time search across song titles, artists, and albums
- **Minimum Duration Filter**: Configurable cutoff to exclude short audio files
- **Dark Mode**: Toggle between light and dark themes
- **Persistent Cache**: Fast startup by caching song metadata
- **Progress Bar**: Visual playback progress with click-to-seek functionality

## Requirements

- Python 3.8+
- pygame
- mutagen (optional, for better metadata extraction)
- pillow (optional, for background images)

## Installation

1. Clone or download this repository
2. Install dependencies:

```bash
pip install pygame mutagen pillow
```

## Project Structure

```
music_player/
├── osu_mp3_browser/          # Main package
│   ├── __init__.py           # Package initialization
│   ├── config.py             # Configuration and constants
│   ├── utils.py              # Utility functions
│   ├── metadata.py           # Audio metadata extraction
│   ├── audio.py              # Audio playback wrapper
│   └── ui.py                 # GUI implementation
└── main.py                   # Application entry point
```

## Controls

### Playback

- **Play**: Start playing selected song
- **Pause/Resume**: Toggle pause state
- **Stop**: Stop playback
- **Progress Bar**: Click or drag to seek to specific position

### Play Modes

- **Sequential**: Play songs in order
- **Loop**: Repeat current track
- **Shuffle**: Play random songs from entire library

### Other

- **Browse**: Change songs directory
- **Scan Now**: Manually trigger directory rescan
- **Search**: Filter songs by title, artist, or album
- **Min length**: Set minimum duration threshold (in seconds)
- **Dark Mode**: Toggle between light and dark themes
- **Escape**: Toggle fullscreen/windowed mode

## Configuration

The application automatically detects your osu! Songs directory at:
```
C:\Users\<username>\AppData\Local\osu!\Songs
```

You can change this by clicking the "Browse..." button.

### Minimum Duration

By default, songs shorter than 30 seconds are excluded. You can adjust this in the UI or by modifying `MIN_DURATION_SECONDS` in `config.py`.

### Cache Location

Song metadata is cached at:
```
~/.osu_mp3_browser_cache.json
```

This speeds up subsequent launches by avoiding re-scanning all files.

## Keyboard Shortcuts

- **Escape**: Toggle between maximized and windowed mode
- **Enter** (in min length field): Apply new duration filter
- **Double-click** (on song): Play selected song

## Troubleshooting

### Audio doesn't play

- Ensure pygame is installed: `pip install pygame`
- Check that your audio files are in supported formats (.mp3, .ogg)

### No songs appear

- Verify your osu! Songs directory exists and contains beatmaps
- Click "Scan Now" to manually trigger a rescan
- Check the console for error messages

### Thumbnails don't show

- Install Pillow: `pip install pillow`
- Ensure .osu files contain valid background image references

### Metadata is missing

- Install mutagen for better tag extraction: `pip install mutagen`
- Some metadata may be parsed from folder names as fallback

## License

This project is provided as-is for personal use.

## Credits

Built with:

- [pygame](https://www.pygame.org/) - Audio playback
- [mutagen](https://mutagen.readthedocs.io/) - Metadata extraction
- [Pillow](https://python-pillow.org/) - Image processing
- [tkinter](https://docs.python.org/3/library/tkinter.html) - GUI framework
- [Inno](https://jrsoftware.org/isinfo.php) - installer
