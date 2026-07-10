"""
Synergy Presentation Converter + Folder Watcher (Python)
- Watches C:\\SynergyPPTs\\Inputs for new/changed .ppt/.pptx files
- Converts each deck onto the SynergyMaster template
- Saves branded output to C:\\SynergyPPTs\\Outputs
Run:  py synergy_convert.py          (watch mode, keeps running)
      py synergy_convert.py --once   (single conversion run, then exit)
"""

import os
import re
import sys
import time
import logging
from pathlib import Path

import pythoncom
import win32com.client

# ---------------- configuration ----------------
BASE_DIR   = Path(r"C:\SynergyPPTs")
INPUTS     = BASE_DIR / "Inputs"
OUTPUTS    = BASE_DIR / "Outputs"
LOG_FILE   = BASE_DIR / "convert_log.txt"
TEMPLATES  = ["SynergyMaster.potx", "Presentation.potx", "Presentation.pptx"]

FORCE_TEMPLATE_FONTS = True
REMOVE_HEADING_BARS  = True

HEAD_FONT  = "Gilroy Bold"
BODY_FONT  = "Gilroy"
NAVY_BGR   = 0x633836          # BGR of #363863
BLACK_BGR  = 0x000000
WHITE_BGR  = 0xFFFFFF

POLL_SECONDS = 5

# PowerPoint constants
PP_SAVE_AS_PPTX   = 24
MSO_GROUP         = 6
MSO_PLACEHOLDER   = 14
PH_TITLE, PH_CENTER_TITLE = 1, 3
PH_CHROME         = (13, 15, 16)      # slide number, footer, date
SKIP_TEXT         = "Private and confidential"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"),
              logging.StreamHandler()],
)
log = logging.getLogger("synergy")


class PowerPointDead(RuntimeError):
    """Raised when the COM connection to PowerPoint is lost."""


def is_rpc_dead(err) -> bool:
    s = str(err)
    return "RPC server is unavailable" in s or "-2147023174" in s


# ---------------- helpers ----------------
def find_template():
    for name in TEMPLATES:
        p = BASE_DIR / name
        if p.exists():
            return p
    return None


def shape_text(shape) -> str:
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return (shape.TextFrame.TextRange.Text or "").strip()
    except Exception:
        pass
    return ""


def para_font_size(shape, default=12.0) -> float:
    try:
        return float(shape.TextFrame.TextRange.Paragraphs(1).Font.Size)
    except Exception:
        return default


def strip_numbering(text: str) -> str:
    return re.sub(r"^\s*\d+[\.\)]\s*", "", text)


def set_template_fonts(shape, is_title: bool):
    """Recursively force template fonts/colours on a shape tree."""
    try:
        if shape.Type == MSO_GROUP:
            for sub in shape.GroupItems:
                set_template_fonts(sub, False)
            return
        if shape.HasTextFrame and shape.TextFrame.HasText:
            f = shape.TextFrame.TextRange.Font
            if is_title:
                f.Name, f.Bold, f.Color.RGB = HEAD_FONT, True, NAVY_BGR
            else:
                f.Name, f.Color.RGB = BODY_FONT, BLACK_BGR
        if shape.HasTable:
            tbl = shape.Table
            for r in range(1, tbl.Rows.Count + 1):
                for c in range(1, tbl.Columns.Count + 1):
                    cf = tbl.Cell(r, c).Shape.TextFrame.TextRange.Font
                    cf.Name, cf.Color.RGB = BODY_FONT, BLACK_BGR
    except Exception:
        pass


def new_powerpoint():
    """Start a PRIVATE PowerPoint instance (DispatchEx) so closing a user's
    PowerPoint window never kills the converter's connection."""
    return win32com.client.DispatchEx("PowerPoint.Application")


# ---------------- core conversion ----------------
def convert_file(app, template: Path, input_file: Path, output_file: Path) -> bool:
    doc = None
    try:
        if output_file.exists():
            try:
                output_file.unlink()
            except PermissionError:
                log.error(f"  SKIPPED (output open in PowerPoint): {output_file.name}")
                return False
        lock = output_file.with_name("~$" + output_file.name)
        if lock.exists():
            try: lock.unlink()
            except Exception: pass

        # fresh untitled copy FROM template (Untitled=True -> template untouched)
        doc = app.Presentations.Open(str(template), False, True, False)
        sw, sh = doc.PageSetup.SlideWidth, doc.PageSetup.SlideHeight

        # capture branded layouts BEFORE anything foreign enters
        brand = {}
        for i in range(1, doc.SlideMaster.CustomLayouts.Count + 1):
            l = doc.SlideMaster.CustomLayouts.Item(i)
            brand.setdefault(l.Name, l)
        brand_count = doc.SlideMaster.CustomLayouts.Count

        while doc.Slides.Count > 0:
            doc.Slides.Item(1).Delete()

        inserted = doc.Slides.InsertFromFile(str(input_file), 0)
        log.info(f"  inserted {inserted} slide(s)")

        # --- remap: slide 1 ALWAYS 'Title Slide'; others by base name ---
        for slide in doc.Slides:
            target = None
            if slide.SlideIndex == 1 and "Title Slide" in brand:
                target = brand["Title Slide"]
            else:
                base = re.sub(r"^\d+_", "", slide.CustomLayout.Name)
                if base in brand and base != "Title Slide":
                    target = brand[base]
                elif "Title and Content" in brand:
                    target = brand["Title and Content"]
            if target is not None:
                try: slide.CustomLayout = target
                except Exception: pass

        # --- REBUILD Title Slide slides: text only, branded design ---
        for slide in doc.Slides:
            if slide.CustomLayout.Name != "Title Slide":
                continue
            try:
                cands = []
                for shape in list(slide.Shapes):
                    txt = shape_text(shape)
                    if not txt or txt == SKIP_TEXT:
                        continue
                    if shape.Rotation != 0:
                        continue
                    cands.append((para_font_size(shape, 18), shape.Top, txt))
                if not cands:
                    continue
                cands.sort(key=lambda t: (-t[0], t[1]))
                title_text = strip_numbering(cands[0][2])
                sub_lines  = [c[2] for c in sorted(cands[1:], key=lambda t: t[1])]

                for shape in list(slide.Shapes):
                    try: shape.Delete()
                    except Exception: pass

                slide.Shapes.AddTitle()
                slide.Shapes.Title.TextFrame.TextRange.Text = title_text

                if sub_lines:
                    t = slide.Shapes.Title
                    sub = slide.Shapes.AddTextbox(1, t.Left, t.Top + t.Height + 6,
                                                  t.Width, 80)
                    sub.TextFrame.TextRange.Text = "\r".join(sub_lines)
                    sf = sub.TextFrame.TextRange.Font
                    sf.Name, sf.Size, sf.Color.RGB = BODY_FONT, 16, WHITE_BGR
            except Exception as e:
                if is_rpc_dead(e):
                    raise
                log.warning(f"  title-slide rebuild issue: {e}")

        # --- remove old heading bars (wide, short, top, textless) ---
        if REMOVE_HEADING_BARS:
            for slide in doc.Slides:
                if slide.CustomLayout.Name == "Title Slide":
                    continue
                for shape in list(slide.Shapes):
                    try:
                        if shape.Type == MSO_PLACEHOLDER:
                            continue
                        if shape_text(shape):
                            continue
                        if (shape.Width > 0.9 * sw and
                                shape.Height < 0.18 * sh and
                                shape.Top < 0.12 * sh):
                            shape.Delete()
                    except Exception:
                        pass

        # --- promote LARGEST-font top-zone text into the title ---
        for slide in doc.Slides:
            if slide.CustomLayout.Name == "Title Slide":
                continue
            try:
                if not slide.Shapes.HasTitle:
                    try: slide.Shapes.AddTitle()
                    except Exception: continue
                title = slide.Shapes.Title
                if shape_text(title):
                    continue

                cands = []
                for shape in list(slide.Shapes):
                    if shape.Id == title.Id:
                        continue
                    txt = shape_text(shape)
                    if not txt or txt == SKIP_TEXT:
                        continue
                    if shape.Rotation != 0:
                        continue
                    if (shape.Type == MSO_PLACEHOLDER and
                            shape.PlaceholderFormat.Type in PH_CHROME):
                        continue
                    if shape.Top > 0.35 * sh:
                        continue
                    cands.append((para_font_size(shape), shape.Top, shape))
                if not cands:
                    continue
                cands.sort(key=lambda t: (-t[0], t[1]))
                candidate = cands[0][2]

                para = candidate.TextFrame.TextRange.Paragraphs(1)
                line = (para.Text or "").strip()
                if not line:
                    continue
                title.TextFrame.TextRange.Text = strip_numbering(line)
                try:
                    title.Fill.Visible = 0
                    title.Line.Visible = 0
                except Exception:
                    pass

                # snap to layout title position + branded typography
                for lps in slide.CustomLayout.Shapes:
                    if (lps.Type == MSO_PLACEHOLDER and
                            lps.PlaceholderFormat.Type in (PH_TITLE, PH_CENTER_TITLE)):
                        title.Left, title.Top = lps.Left, lps.Top
                        title.Width, title.Height = lps.Width, lps.Height
                        break
                tf = title.TextFrame.TextRange.Font
                tf.Name, tf.Bold, tf.Color.RGB = HEAD_FONT, True, NAVY_BGR

                para.Delete()
                if not shape_text(candidate):
                    candidate.Delete()
                else:
                    try:
                        candidate.Fill.Visible = 0
                        candidate.Line.Visible = 0
                    except Exception:
                        pass

                # clear title-area leftovers: textless decor + stale short headings
                tT, tB = title.Top, title.Top + title.Height
                tL, tR = title.Left, title.Left + title.Width
                for shape in list(slide.Shapes):
                    try:
                        if shape.Id == title.Id or shape.Type == MSO_PLACEHOLDER:
                            continue
                        overlaps = (shape.Top < tB and shape.Top + shape.Height > tT and
                                    shape.Left < tR and shape.Left + shape.Width > tL)
                        if not overlaps:
                            continue
                        txt = shape_text(shape)
                        if (not txt and shape.Type in (1, 5)) or \
                           (txt and len(txt) <= 60 and shape.Height < 0.15 * sh):
                            shape.Delete()
                    except Exception:
                        pass
            except Exception as e:
                if is_rpc_dead(e):
                    raise
                log.warning(f"  title promotion issue on slide {slide.SlideIndex}: {e}")

        # --- delete empty placeholders (titles + 'Click to add text') ---
        for slide in doc.Slides:
            for shape in list(slide.Shapes):
                try:
                    if shape.Type != MSO_PLACEHOLDER:
                        continue
                    if shape.HasTable or shape.HasChart:
                        continue
                    if shape.HasTextFrame and not shape_text(shape):
                        shape.Delete()
                except Exception:
                    pass

        # --- FORCE template fonts & colours on content slides ---
        if FORCE_TEMPLATE_FONTS:
            for slide in doc.Slides:
                if slide.CustomLayout.Name == "Title Slide":
                    continue
                for shape in list(slide.Shapes):
                    is_title = False
                    try:
                        if shape.Type == MSO_PLACEHOLDER:
                            is_title = shape.PlaceholderFormat.Type in (PH_TITLE,
                                                                        PH_CENTER_TITLE)
                    except Exception:
                        pass
                    set_template_fonts(shape, is_title)

        # --- purge imported layouts beyond the branded set ---
        for i in range(doc.SlideMaster.CustomLayouts.Count, brand_count, -1):
            try: doc.SlideMaster.CustomLayouts.Item(i).Delete()
            except Exception: pass

        doc.SaveAs(str(output_file), PP_SAVE_AS_PPTX)
        log.info(f"  saved: {output_file.name}")
        return True

    except Exception as e:
        if is_rpc_dead(e):
            log.error(f"  FAILED {input_file.name}: PowerPoint connection lost — restarting")
            raise PowerPointDead() from e
        log.error(f"  FAILED {input_file.name}: {e}")
        return False
    finally:
        if doc is not None:
            try:
                doc.Saved = True
                doc.Close()
            except Exception:
                pass


def convert_all() -> None:
    template = find_template()
    if template is None:
        log.error(f"No template found in {BASE_DIR} ({', '.join(TEMPLATES)})")
        return
    OUTPUTS.mkdir(exist_ok=True)

    files = [p for p in INPUTS.iterdir()
             if p.suffix.lower() in (".ppt", ".pptx")
             and not p.name.startswith("~$")]
    if not files:
        log.info("No input files.")
        return

    log.info(f"Template: {template.name} | Files: {len(files)}")
    pythoncom.CoInitialize()
    app = None
    ok = fail = 0
    try:
        app = new_powerpoint()
        for f in sorted(files):
            log.info(f"Processing: {f.name}")
            out = OUTPUTS / (f.stem + ".pptx")
            try:
                if convert_file(app, template, f, out):
                    ok += 1
                else:
                    fail += 1
            except PowerPointDead:
                # PowerPoint died - restart a fresh instance and retry this file once
                try:
                    app.Quit()
                except Exception:
                    pass
                try:
                    app = new_powerpoint()
                except Exception as e:
                    log.error(f"Could not restart PowerPoint: {e}")
                    fail += 1
                    break
                log.info(f"  retrying: {f.name}")
                try:
                    if convert_file(app, template, f, out):
                        ok += 1
                    else:
                        fail += 1
                except PowerPointDead:
                    log.error(f"  FAILED {f.name}: PowerPoint unstable — aborting batch")
                    fail += 1
                    break
    finally:
        if app is not None:
            try: app.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()
    log.info(f"Done. Succeeded: {ok} | Failed: {fail}")


# ---------------- folder watcher (polling) ----------------
def file_ready(path: Path, checks=3, gap=2) -> bool:
    """File is ready when its size stops changing and it opens exclusively."""
    last = -1
    for _ in range(checks):
        try:
            size = path.stat().st_size
        except OSError:
            return False
        if size == last:
            try:
                with open(path, "rb"):
                    return True
            except OSError:
                pass
        last = size
        time.sleep(gap)
    return False


def snapshot() -> dict:
    return {p.name: p.stat().st_mtime
            for p in INPUTS.iterdir()
            if p.suffix.lower() in (".ppt", ".pptx")
            and not p.name.startswith("~$")}


def watch() -> None:
    log.info(f"Watcher started. Monitoring: {INPUTS} (every {POLL_SECONDS}s)")
    seen = snapshot()
    while True:
        time.sleep(POLL_SECONDS)
        try:
            now = snapshot()
        except FileNotFoundError:
            log.error("Inputs folder missing!")
            continue
        new_or_changed = [n for n, m in now.items()
                          if n not in seen or m > seen[n]]
        if new_or_changed:
            log.info(f"Detected: {', '.join(new_or_changed)}")
            for name in new_or_changed:
                file_ready(INPUTS / name)
            try:
                convert_all()
            except Exception as e:
                log.error(f"Batch error: {e}")
            seen = snapshot()
        else:
            seen = now


if __name__ == "__main__":
    if "--once" in sys.argv:
        convert_all()
    else:
        watch()
