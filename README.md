# Keynote-PowerPoint via PDF

## The Issue

I use Keynote to make my presentations, but when I can't present off of my own laptop, I have to either:
  - Export to PowerPoint
  - Export to PDF, one-slide-per-build.

I don't like to export to PowerPoint straight from because I use a lot of custom fonts and usually whoever I send the slides to doesn't have the fonts and so things look terrible. Also I use a lot of animations/builds with grouped objects, and those sometimes get mangled in the conversion.

So I usually export to PDF, with one slide per build. This works great, but has two problems:

  - Whatever computer we're presenting off of has to have its PDF reader set correctly (in "single-page" mode, full screen) and that's caused issues in the past.
  - More than once, the AV people for the session want to combine everybody's slides into a giant meta-PowerPoint; my slides, being in PDF, don't fit in, and that causes delays and confusion.
    - And there's an even worse scenario that happened last year: they tried to use Acrobat to convert the PDF to a Powerpoint, or something, and it got everything _really_ FUBARed.

## The Solution

`pdf_to_pptx.py` takes the PDF that Keynote already produces and wraps each
page in a proper PPTX slide, with the page rendered as a full-bleed JPEG image.
The result looks identical to the PDF — fonts, layouts, build order — but opens
natively in PowerPoint. Optionally, it can reach back into the original Keynote
file and pull your speaker notes across too.

---

## Requirements

- macOS (the speaker-notes feature uses AppleScript to talk to Keynote;
  everything else works cross-platform)
- Python 3.13+ managed by [uv](https://github.com/astral-sh/uv)
- Apple Keynote (only needed for the `--with_notes` feature)

Install dependencies:

```sh
uv sync
```

---

## Basic Usage

### Single file

```sh
uv run python pdf_to_pptx.py "My Talk.pdf"
```

Output is written to `My Talk.pptx` in the same directory as the input.

### Specify an output path

```sh
uv run python pdf_to_pptx.py "My Talk.pdf" --output ~/Desktop/MyTalk.pptx
```

### Batch mode — convert a whole directory

```sh
uv run python pdf_to_pptx.py ./pdfs --batch true
```

Each `.pdf` file in the directory gets its own `.pptx` written alongside it.
To send all output to a different directory:

```sh
uv run python pdf_to_pptx.py ./pdfs --batch true --output ./pptx_output
```

---

## Speaker Notes

With `--with_notes true`, the script exports a temporary PPTX from the
matching Keynote file, reads your speaker notes out of it, and embeds them into
the output PPTX. The temporary file is deleted immediately afterwards.

```sh
uv run python pdf_to_pptx.py "My Talk.pdf" --with_notes true
```

If you have exactly one Keynote document open, it is used automatically. If you
have several open, you are shown a numbered menu and asked to pick one — the
script will try to highlight a sensible default based on filename similarity.

To skip the prompt entirely, point `--keynote` at the file directly:

```sh
uv run python pdf_to_pptx.py "My Talk.pdf" --with_notes true --keynote ~/Decks/MyTalk.key
```

In batch mode, `--keynote` is ignored and you are prompted once per file.

### How the notes mapping works

Keynote's "one slide per build" PDF export produces one page per *build step*,
while a Keynote file has one slide per *actual slide* (with builds encoded as
animations). The script bridges this gap by:

1. Exporting a temporary PPTX from Keynote (which also has one slide per
   Keynote slide, with builds as OOXML animation sequences).
2. Reading the animation timing tree of each visible slide to count how many
   click-advance build steps it contains.
3. Computing `pages_for_this_slide = build_steps + 1`.
4. Assigning the same speaker notes to *every* PDF page that belongs to a
   given slide — so your notes show up in Presenter View throughout all the
   build steps of that slide, not just on the last one.

### Notes mismatch warning

If the page count predicted from the animation data doesn't match the actual
PDF page count, the summary table will show a yellow **mismatch** warning.
Notes are still assigned on a best-effort basis to whatever slides do line up.

The most common cause is **hidden/skipped slides**: both the PDF export and the
PPTX export skip them, so they should normally agree. But if slides were hidden
*after* the PDF was exported, or if the PDF was exported with different
settings, the counts can diverge. Re-exporting the PDF from Keynote and
re-running the script should resolve it.

---

## Options

| Flag | Type | Default | Description |
|---|---|---|---|
| `input` | `Path` | *(required)* | PDF file to convert, or directory in batch mode |
| `--output` | `Path` | same dir as input | Output `.pptx` path (single mode) or directory (batch mode) |
| `--dpi` | `int` | `150` | Render resolution. 150 is sharp on a projector; raise to 200-300 for print-quality output (larger files) |
| `--jpeg_quality` | `int` | `85` | JPEG compression quality, 1-95. 85 is a good balance; lower = smaller files, higher = better quality |
| `--batch` | `bool` | `false` | Treat `input` as a directory and convert all `.pdf` files in it |
| `--with_notes` | `bool` | `false` | Extract and embed speaker notes from the matching Keynote file |
| `--keynote` | `Path` | *(picker)* | Explicit path to the `.key` file to pull notes from; skips the interactive prompt |

---

## File Size

File size is primarily a function of DPI, JPEG quality, and the number of
slides. Some rough benchmarks at the default settings (150 DPI, quality 85):

- A typical 16:9 slide renders to roughly **100-200 KB** as a JPEG.
- A 100-page PDF (e.g. ~30 slides with a few builds each) produces roughly a **15-25 MB** PPTX.
- A 200-page PDF produces roughly a **30-50 MB** PPTX.

If size is a concern, try dropping `--jpeg_quality` to `75` or `--dpi` to
`120` first — quality is rarely perceptible on a projector.

---

## Limitations

- **Slide images are rasterized.** Each slide is a JPEG, not vector art. At
  150 DPI this is imperceptible on a projector or screen, but if you zoom in
  very close in PowerPoint you will see pixels. This is an inherent limitation
  of the PPTX format — there is no standard way to embed a PDF page as live
  vector content.

- **No real animations.** The output PPTX advances slide-by-slide; the
  per-build-step sequencing is encoded in slide order, not as PowerPoint
  animations. This is the point — it mirrors exactly what the PDF would do.

- **Speaker notes require macOS and Keynote.** The `--with_notes` feature uses
  AppleScript and therefore only works on macOS with Keynote installed. The
  base PDF-to-PPTX conversion works anywhere.

- **Speaker notes formatting is partially preserved.** Paragraph structure
  (line breaks, bullets) is preserved. Rich inline formatting within a
  paragraph (bold, italic, font size changes) is not carried across.

- **One slide size for the whole deck.** Slide dimensions are taken from the
  first PDF page and applied to all slides. If your PDF somehow has pages of
  different sizes, later pages will be stretched or letterboxed.

- **Batch mode + `--with_notes` is interactive.** You will be prompted once
  per file to identify the source Keynote document. This is intentional — batch
  mode is designed for converting unrelated PDFs that each have their own
  Keynote source.

---

## Troubleshooting

### "No Keynote documents are currently open"

The script talks to the running Keynote app via AppleScript. Make sure Keynote
is open and the relevant `.key` file is loaded before running with
`--with_notes true`. Alternatively, pass `--keynote /path/to/file.key` to have
the script open the file automatically.

### Notes mismatch warning in the summary

See [Notes mismatch warning](#notes-mismatch-warning) above. The most likely
fix is to re-export the PDF from Keynote and re-run the script so both exports
are in sync.

### The output PPTX looks blurry

Try increasing `--dpi` to `200` or `300`. Bear in mind this will increase file
size roughly as the square of the DPI ratio (200 DPI produces 4x the pixels of
100 DPI).

### AppleScript permission error / osascript fails

On macOS, Terminal (or whichever shell you use) needs Automation permission to
control Keynote. If you have never granted this, macOS should prompt you
automatically the first time. If it does not, go to **System Settings ->
Privacy & Security -> Automation** and make sure your terminal app has
permission to control Keynote.

### The script produces a 0-byte or corrupted PPTX

This usually means `python-pptx` failed to save, most likely due to a full
disk or a permissions problem on the output directory. Check that you have
write access to the output location and that there is sufficient disk space.
