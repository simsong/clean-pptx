# clean-pptx

Clean and re-render PowerPoint decks.

## Usage

```bash
python clean_pptx.py /path/to/deck.pptx
```

This creates:

- `/path/to/deck-cleaned.pptx`
- `/path/to/deck-cleaned.pdf`

The cleaned PPTX is rebuilt slide-by-slide using `python-pptx`, text fonts are migrated to standard fonts, and images are re-encoded as JPEG quality 90.

The PDF rendering step uses AppleScript to remote-control Microsoft PowerPoint on macOS.

If you only want the cleaned PPTX:

```bash
python clean_pptx.py /path/to/deck.pptx --skip-pdf
```

## Requirements

- Python 3
- `python-pptx`
- `Pillow`
- macOS + Microsoft PowerPoint (for PDF rendering)
