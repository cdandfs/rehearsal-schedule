"""
make-student-schedules.py
Generates one landscape PDF per student by joining three input sources:

  input.xlsx  "classes"   tab  →  class metadata (teacher, day, costume…)
  input.xlsx  "rehearsals" tab →  all rehearsal / performance calls
  class-rosters.xlsx  "roster" tab  →  which students are in which class

The join logic is:
    roster.class_name  →  classes.class_name  (get teacher, day, costume)
    roster.class_name  →  rehearsals.class_name  (get calls for that class)

One PDF is written per student containing:
    • their enrolled classes (with teacher, day/time, costume)
    • every rehearsal / performance call across all their classes

Dependencies:
    pip install weasyprint jinja2 pandas openpyxl

Usage:
    python make-student-schedules.py
    python make-student-schedules.py --schedule input.xlsx \\
                                     --rosters  class-rosters.xlsx \\
                                     --output-dir output/ \\
                                     --templates templates/
"""

import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader

# ── PDF backend ────────────────────────────────────────────────────────────
try:
    import weasyprint
    def _html_to_pdf(html_string: str, dest_path: str) -> None:
        weasyprint.HTML(string=html_string).write_pdf(dest_path)
    PDF_BACKEND = "weasyprint"
except ImportError:
    try:
        import pdfkit
        _PDFKIT_OPTIONS = {
            "page-size":    "Letter",
            "orientation":  "Landscape",
            "margin-top":    "0.45in",
            "margin-bottom": "0.45in",
            "margin-left":   "0.55in",
            "margin-right":  "0.55in",
            "enable-local-file-access": "",
            "enable-external-links": "",
            "quiet": "",
        }
        def _html_to_pdf(html_string: str, dest_path: str) -> None:
            pdfkit.from_string(html_string, dest_path, options=_PDFKIT_OPTIONS)
        PDF_BACKEND = "pdfkit"
    except ImportError:
        raise SystemExit(
            "No PDF backend found. Install weasyprint:\n"
            "    pip install weasyprint"
        )

# ── File & sheet names ─────────────────────────────────────────────────────
SCHEDULE_FILE  = "input.xlsx"
ROSTERS_FILE   = "class-rosters.xlsx"

CLASSES_SHEET   = "classes"
REHEARSALS_SHEET = "rehearsals"
ROSTER_SHEET    = "roster"

# ── Column names — input.xlsx "classes" ───────────────────────────────────
CLS_CLASS     = "class_name"
CLS_TEACHER   = "teacher"
CLS_ASSISTANT = "assistant"
CLS_DAY       = "day_of_week"
CLS_TIME      = "time_of_day"
CLS_COSTUME   = "costume"

# ── Column names — input.xlsx "rehearsals" ────────────────────────────────
RH_NAME       = "name"           # event type, e.g. "Technical Rehearsal"
RH_DATE       = "date"           # e.g. "Sun, May 11"
RH_LOCATION   = "location"
RH_CLASS      = "class_name"
RH_DANCE      = "dance_name"
RH_START      = "start_time"
RH_END        = "end_time"
RH_ARRIVAL    = "arrival_time"
RH_INFO       = "information"    # notes / extra info
RH_URL        = "url"

# ── Column names — class-rosters.xlsx "roster" ────────────────────────────
RS_CLASS             = "class_name"
RS_STUDENT           = "student"
RS_DRESSING_KIDS     = "dressing_room_kids"
RS_DRESSING_BALLET   = "dressing_room_ballet"
RS_DRESSING_TEENADULT = "dressing_room_teenadult"
RS_DRESSING_MATINEE  = "dressing_room_matinee"

# ── Studio / show metadata ─────────────────────────────────────────────────
STUDIO_NAME      = "Creative Dance & Fitness Studio"
PERFORMANCE_NAME = "Spring Performance 2026"


# ── Helpers ────────────────────────────────────────────────────────────────

def _str(val) -> str:
    """Clean string, or '' for NaN / None."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _fmt_time(val) -> str:
    """Normalise a time value to H:MM AM/PM."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if hasattr(val, "strftime"):
        return val.strftime("%-I:%M %p")
    s = str(val).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M%p"):
        try:
            return datetime.strptime(s, fmt).strftime("%-I:%M %p")
        except ValueError:
            pass
    return s


def _badge_class(event_name: str) -> str:
    n = event_name.lower()
    if "studio" in n:
        return "pill-studio"
    if "technical" in n:
        return "pill-tech"
    if "dress" in n:
        return "pill-dress"
    if "performance" in n:
        return "pill-perf"
    return "pill-studio"


def _short_event(event_name: str) -> str:
    mapping = {
        "studio rehearsal":    "Studio Rehearsal",
        "technical rehearsal": "Technical Rehearsal",
        "dress rehearsal":     "Dress Rehearsal",
        "performance":         "Performance",
    }
    return mapping.get(event_name.strip().lower(), event_name.strip())


def _safe_filename(name: str) -> str:
    return re.sub(r"[^\w\-]", "_", name.strip())


def _parse_date(raw) -> tuple[str, str]:
    """
    Return (day_of_week_abbrev, "Month Day") from several input forms:
      - datetime / Timestamp  →  ("SUN", "May 11")
      - "Sun, May 11"         →  ("SUN", "May 11")
      - "2025-05-11"          →  ("SUN", "May 11")
      - "May 11"              →  ("",    "May 11")
    """
    # Excel / pandas may give us a datetime object directly
    if hasattr(raw, "strftime"):
        return raw.strftime("%a").upper(), raw.strftime("%b %-d")

    s = str(raw).strip()

    # Try parsing as an ISO date string
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%a").upper(), dt.strftime("%b %-d")
        except ValueError:
            pass

    # "Sun, May 11" — split on first comma
    if "," in s:
        day, rest = s.split(",", 1)
        rest = rest.strip()
        # Try to parse rest as a full date to normalise it
        for fmt in ("%B %d", "%b %d", "%B %d %Y", "%b %d %Y"):
            try:
                dt = datetime.strptime(rest, fmt)
                return day.strip().upper(), dt.strftime("%b %-d")
            except ValueError:
                pass
        return day.strip().upper(), rest

    return "", s


# ── Data loading ───────────────────────────────────────────────────────────

def load_classes(schedule_path: str) -> dict[str, dict]:
    """Return {class_name: class_metadata_dict}."""
    df = pd.read_excel(schedule_path, sheet_name=CLASSES_SHEET)
    result: dict[str, dict] = {}
    for _, row in df.iterrows():
        name = _str(row.get(CLS_CLASS))
        if not name:
            continue
        result[name] = {
            "name":      name,
            "teacher":   _str(row.get(CLS_TEACHER)),
            "assistant": _str(row.get(CLS_ASSISTANT)),
            "day":       _str(row.get(CLS_DAY)),
            "time":      _fmt_time(row.get(CLS_TIME)),
            "costume":   _str(row.get(CLS_COSTUME)),
        }
    return result


def load_rehearsals(schedule_path: str) -> dict[str, list[dict]]:
    """Return {class_name: [rehearsal_dict, ...]}."""
    df = pd.read_excel(schedule_path, sheet_name=REHEARSALS_SHEET)

    result: dict[str, list[dict]] = {}
    for _, row in df.iterrows():
        class_name = _str(row.get(RH_CLASS))
        if not class_name:
            continue

        event_name = _str(row.get(RH_NAME))
        raw_date   = row.get(RH_DATE)
        day_of_week, date_display = _parse_date(raw_date)

        # Build a proper datetime sort key so May 2 < May 10
        sort_date = pd.NaT
        if hasattr(raw_date, "strftime"):
            sort_date = raw_date
        else:
            for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%B %d", "%b %d"):
                try:
                    sort_date = datetime.strptime(str(raw_date).split(",")[-1].strip(), fmt)
                    break
                except ValueError:
                    pass

        result.setdefault(class_name, []).append({
            "event_type":   _short_event(event_name),
            "badge_class":  _badge_class(event_name),
            "date":         date_display,
            "day_of_week":  day_of_week,
            "location":     _str(row.get(RH_LOCATION)),
            "dance_name":   _str(row.get(RH_DANCE)),
            "start_time":   _fmt_time(row.get(RH_START)),
            "end_time":     _fmt_time(row.get(RH_END)),
            "arrival_time": _fmt_time(row.get(RH_ARRIVAL)),
            "information":  _str(row.get(RH_INFO)),
            "url":          _str(row.get(RH_URL)),
            "_sort_date":   sort_date,
            "_sort_start":  row.get(RH_START),
        })
    return result


def load_rosters(rosters_path: str) -> dict[str, dict]:
    """Return {student_name: {"classes": [...], "dressing_rooms": {...}}}.

    dressing_rooms contains:
        kids      - dressing_room_kids value
        ballet    — dressing_room_ballet value
        teenadult — dressing_room_teenadult value
        matinee   — dressing_room_matinee value
        differ    — True if any non-empty values differ from each other

    A student may appear on multiple rows; the first non-empty value for each
    room column is used.
    """
    df = pd.read_excel(rosters_path, sheet_name=ROSTER_SHEET)
    result: dict[str, dict] = {}
    for _, row in df.iterrows():
        student    = _str(row.get(RS_STUDENT))
        class_name = _str(row.get(RS_CLASS))
        if not student or not class_name:
            continue
        entry = result.setdefault(student, {
            "classes": [],
            "dressing_rooms": {"kids": "", "ballet": "", "teenadult": "", "matinee": ""},
        })
        entry["classes"].append(class_name)
        rooms = entry["dressing_rooms"]
        if not rooms["kids"]:
            rooms["kids"]      = _str(row.get(RS_DRESSING_KIDS))
        if not rooms["ballet"]:
            rooms["ballet"]    = _str(row.get(RS_DRESSING_BALLET))
        if not rooms["teenadult"]:
            rooms["teenadult"] = _str(row.get(RS_DRESSING_TEENADULT))
        if not rooms["matinee"]:
            rooms["matinee"]   = _str(row.get(RS_DRESSING_MATINEE))

    # Compute whether the non-empty room values differ
    for entry in result.values():
        rooms = entry["dressing_rooms"]
        non_empty = {v for v in rooms.values() if v}
        rooms["differ"] = len(non_empty) > 1

    return result


# ── Per-student data assembly ──────────────────────────────────────────────

def build_student_data(
    student: str,
    enrolled_classes: list[str],
    all_classes: dict[str, dict],
    all_rehearsals: dict[str, list[dict]],
) -> tuple[list[dict], list[dict]]:
    """
    Returns (classes_list, rehearsals_list) for one student.

    classes_list     — one entry per enrolled class
    rehearsals_list  — all calls across all enrolled classes, in date order
    """
    classes = []
    rehearsals = []

    for class_name in enrolled_classes:
        if class_name in all_classes:
            classes.append(all_classes[class_name])
        for r in all_rehearsals.get(class_name, []):
            rehearsals.append({**r, "class_name": class_name})

    # Sort by actual date then start time, handling NaT safely
    def _sort_key(r):
        d = r["_sort_date"]
        s = r["_sort_start"]
        # Convert to a comparable value; put NaT/None last
        d_val = d if pd.notna(d) else datetime.max
        s_val = s if (s is not None and pd.notna(s)) else datetime.max
        return (d_val, s_val)

    rehearsals.sort(key=_sort_key)

    # Strip internal keys before handing to template
    for r in rehearsals:
        r.pop("_sort_date", None)
        r.pop("_sort_start", None)

    return classes, rehearsals


# ── Main ───────────────────────────────────────────────────────────────────

def generate_schedules(
    schedule_path: str,
    rosters_path:  str,
    output_dir:    str,
    template_dir:  str,
) -> None:
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    env      = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template("schedule.html")

    all_classes    = load_classes(schedule_path)
    all_rehearsals = load_rehearsals(schedule_path)
    rosters        = load_rosters(rosters_path)

    print(f"PDF backend : {PDF_BACKEND}")
    print(f"Classes     : {len(all_classes)}")
    print(f"Students    : {len(rosters)}")
    print(f"Output dir  : {output_path.resolve()}\n")

    for student in sorted(rosters):
        enrolled      = rosters[student]["classes"]
        dressing_rooms = rosters[student]["dressing_rooms"]
        classes, rehearsals = build_student_data(
            student, enrolled, all_classes, all_rehearsals
        )

        html = template.render(
            student_name     = student,
            studio_name      = STUDIO_NAME,
            performance_name = PERFORMANCE_NAME,
            generated_date   = datetime.today().strftime("%B %-d, %Y"),
            classes          = classes,
            rehearsals       = rehearsals,
            dressing_rooms   = dressing_rooms,
        )

        dest = output_path / f"{_safe_filename(student)}.pdf"
        _html_to_pdf(html, str(dest))
        print(f"  ✓  {student:30s}  ({len(rehearsals):2d} calls)  →  {dest.name}")

    print(f"\nDone — {len(rosters)} PDF(s) written to {output_path.resolve()}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate per-student rehearsal PDFs from roster + schedule data."
    )
    parser.add_argument("--schedule",    default=SCHEDULE_FILE, help="Path to input.xlsx")
    parser.add_argument("--rosters",     default=ROSTERS_FILE,  help="Path to class-rosters.xlsx")
    parser.add_argument("--output-dir",  default="student-schedules/", help="Directory to write PDFs")
    parser.add_argument("--templates",   default="templates/",  help="Directory with schedule.html")
    args = parser.parse_args()

    generate_schedules(args.schedule, args.rosters, args.output_dir, args.templates)
