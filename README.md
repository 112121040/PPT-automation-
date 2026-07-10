# PPT Automation — Presentation Template Converter

Automated PowerPoint re-branding pipeline for Windows. Drop any `.ppt` or
`.pptx` file into an `Inputs` folder, and a converted copy — rebuilt on a
corporate master template — appears in `Outputs` within seconds. Fully
hands-off: a background watcher detects new files and triggers conversion
automatically.

## Problem this solves

Organizations accumulate presentations built on many different templates,
themes, and ad-hoc styles. Manually re-theming decks in PowerPoint is slow
and unreliable: applying a template does not remove old design elements,
does not fix text sitting in plain text boxes instead of placeholders, and
does not override direct (hard-coded) font formatting. This tool automates
the full normalization, not just the theme swap.

## How it works

The converter drives a locally installed Microsoft PowerPoint through its
COM automation API (`pywin32`), which guarantees full rendering fidelity.
For each input deck it:

1. Opens a fresh copy of the master template (`.potx`) and inserts all
   slides from the input file (`Slides.InsertFromFile` — no clipboard).
2. Rebuilds the first slide on the template's title layout: harvests the
   title and subtitle text from the original, discards old artwork, and
   re-creates the text through the layout's placeholders so the branded
   design shows through.
3. Remaps every other slide to the template's layouts by name, with a
   sensible fallback layout for unmatched slides.
4. Promotes headings: many source decks fake titles using plain text boxes.
   The largest-font text in the top zone of each slide is moved into the
   real title placeholder, inheriting the template's position and styling.
   Leading numbering (e.g. `3.`) is stripped.
5. Cleans up leftovers: old decorative heading bars, boxes overlapping the
   title area, empty placeholder prompts, and layouts imported from the
   source deck.
6. Optionally forces template typography across all text, tables, and
   grouped shapes, overriding direct formatting from the source deck.
7. Saves the result to `Outputs\<name>.pptx`. Input files are never
   modified.

## Requirements

- Windows with Microsoft PowerPoint installed
- Python 3.10+ (`py --version`)
- pywin32: `py -m pip install pywin32`
- The template's brand fonts installed (text falls back to system defaults
  if missing)

## Folder layout
<BASE_DIR>
├── synergy_convert.py      # converter + folder watcher
├── <MasterTemplate>.potx   # corporate template (configurable search order)
├── Inputs\                 # drop source decks here
├── Outputs\                # converted results appear here
└── convert_log.txt         # run log
