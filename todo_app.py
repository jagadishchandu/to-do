"""Daily Todo - a system-tray resident Kanban-style todo board for today's tasks."""
import json
import queue
import subprocess
import sys
import threading
import webbrowser
import winsound
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox
from xml.sax.saxutils import escape as xml_escape

import customtkinter as ctk
import pystray
from PIL import Image, ImageDraw
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False

APP_DIR = Path(__file__).resolve().parent
TASKS_FILE = APP_DIR / "tasks.json"
HISTORY_FILE = APP_DIR / "history.json"
NOTES_FILE = APP_DIR / "notes.json"
WEEKLY_FILE = APP_DIR / "weekly_tasks.json"
WEEKLY_HISTORY_FILE = APP_DIR / "weekly_history.json"
HALF_YEAR_FILE = APP_DIR / "half_year_tasks.json"
HALF_YEAR_HISTORY_FILE = APP_DIR / "half_year_history.json"
YEARLY_FILE = APP_DIR / "yearly_tasks.json"
YEARLY_HISTORY_FILE = APP_DIR / "yearly_history.json"
MAIL_WATCHLIST_FILE = APP_DIR / "mail_watchlist.json"
PORTALS_FILE = APP_DIR / "portals.json"
PORTAL_BROWSER_HELPER = APP_DIR / "portal_browser_helper.py"
PORTAL_PROFILE_DIR = APP_DIR / "portal_browser_profile"

# ---- Palette --------------------------------------------------------------
BG = "#0f0f17"
CARD = "#1b1b28"
CARD2 = "#20202f"
CARD_DONE = "#161620"
ACCENT = "#7c6cff"
ACCENT_HOVER = "#6a58f0"
DANGER = "#3a2230"
DANGER_TEXT = "#f28b9c"
TEXT_PRIMARY = "#f2f2f7"
TEXT_SECONDARY = "#9a9ab0"
TEXT_MUTED = "#5c5c72"
SUCCESS = "#4ade80"
ENTRY_BG = "#181824"
BORDER = "#2a2a3a"
SHADOW_COLOR = "#05050a"

STATUS_ORDER = ["todo", "in_progress", "done"]
STATUS_LABELS = {"todo": "To Do", "in_progress": "In Progress", "done": "Done"}
COLUMN_ACCENTS = {"todo": "#8b8bea", "in_progress": "#facc15", "done": SUCCESS}

PRIORITY_ORDER = ["high", "medium", "low"]
PRIORITY_LABELS = {"high": "High", "medium": "Medium", "low": "Low"}
PRIORITY_COLORS = {"high": "#ef4444", "medium": "#eab308", "low": "#8a8aa0"}
PRIORITY_TEXT_COLORS = {"high": "#ffffff", "medium": "#2a2200", "low": "#ffffff"}


# ---- Persistence -----------------------------------------------------------
def migrate_task(t: dict) -> dict:
    if "status" not in t:
        t["status"] = "done" if t.get("done") else "todo"
    t.pop("done", None)
    t.setdefault("notes", "")
    t.setdefault("order", 0)
    t.setdefault("priority", "medium")
    if t["priority"] not in PRIORITY_ORDER:
        t["priority"] = "medium"
    t.setdefault("created_at", None)
    t.setdefault("started_at", None)
    t.setdefault("start_time", None)
    t.setdefault("end_time", None)
    t.setdefault("reminded_for_end_time", None)
    t.setdefault("parent_id", None)
    if "completed_at" not in t:
        legacy_date = t.pop("completed_date", None)
        t["completed_at"] = f"{legacy_date}T00:00:00" if legacy_date else None
    else:
        t.pop("completed_date", None)
    return t


def migrate_note(n: dict) -> dict:
    n.setdefault("title", "")
    n.setdefault("text", "")
    n.setdefault("created_at", None)
    n.setdefault("updated_at", n.get("created_at"))
    if "history" not in n:
        history = []
        if n.get("created_at"):
            history.append({"at": n["created_at"], "action": "created"})
        if n.get("updated_at") and n.get("updated_at") != n.get("created_at"):
            history.append({"at": n["updated_at"], "action": "edited"})
        n["history"] = history
    return n


def fmt_12h(iso_str) -> str:
    if not iso_str:
        return "—"
    try:
        dt = datetime.fromisoformat(iso_str)
    except ValueError:
        return "—"
    text = dt.strftime("%I:%M %p")
    return text.lstrip("0") or text


def fmt_datetime_12h(iso_str) -> str:
    if not iso_str:
        return "—"
    try:
        dt = datetime.fromisoformat(iso_str)
    except ValueError:
        return "—"
    return f"{dt.strftime('%b %d, %Y')} · {fmt_12h(iso_str)}"


def fmt_duration(start_iso, end_iso) -> str:
    if not start_iso or not end_iso:
        return "—"
    try:
        start = datetime.fromisoformat(start_iso)
        end = datetime.fromisoformat(end_iso)
    except ValueError:
        return "—"
    total_minutes = int((end - start).total_seconds() // 60)
    if total_minutes < 0:
        return "—"
    days, rem = divmod(total_minutes, 1440)
    hours, minutes = divmod(rem, 60)
    parts = []
    if days:
        parts.append(f"{days}d")
    if hours:
        parts.append(f"{hours}h")
    if minutes or not parts:
        parts.append(f"{minutes}m")
    return " ".join(parts)


def hhmm_to_iso(hhmm: str, base_date=None):
    """Parse an 'HH:MM' 24-hour string into an ISO datetime on base_date (default: today)."""
    hhmm = (hhmm or "").strip()
    if not hhmm:
        return None
    try:
        t = datetime.strptime(hhmm, "%H:%M").time()
    except ValueError:
        return None
    return datetime.combine(base_date or date.today(), t).isoformat()


def iso_to_hhmm(iso_str) -> str:
    if not iso_str:
        return ""
    try:
        return datetime.fromisoformat(iso_str).strftime("%H:%M")
    except ValueError:
        return ""


def hhmm_display(hhmm: str) -> str:
    """'14:30' -> '2:30 PM'"""
    t = datetime.strptime(hhmm, "%H:%M")
    s = t.strftime("%I:%M %p")
    return s.lstrip("0") or s


def generate_time_slots(interval_minutes: int = 15) -> list:
    """HH:MM 24h strings from now (rounded up to the interval) through the end of today."""
    now = datetime.now()
    remainder = now.minute % interval_minutes
    rounded = now.replace(second=0, microsecond=0)
    if remainder or now.second or now.microsecond:
        rounded += timedelta(minutes=interval_minutes - remainder)
    end_of_day = now.replace(hour=23, minute=45, second=0, microsecond=0)
    slots = []
    t = rounded
    while t <= end_of_day:
        slots.append(t.strftime("%H:%M"))
        t += timedelta(minutes=interval_minutes)
    return slots


def build_time_options(current_hhmm: str = "") -> tuple:
    """Return (labels, label_to_hhmm) for a time dropdown: upcoming slots plus the
    task's already-chosen time (even if it's now in the past) so it stays visible."""
    slots = generate_time_slots()
    if current_hhmm and current_hhmm not in slots:
        slots = sorted(slots + [current_hhmm])
    labels = ["No time"] + [hhmm_display(s) for s in slots]
    label_to_hhmm = {"No time": None}
    for s in slots:
        label_to_hhmm[hhmm_display(s)] = s
    return labels, label_to_hhmm


# ---- Longer-horizon period helpers (weekly / half-year / yearly boards) ---
def week_key(d: date) -> str:
    monday = d - timedelta(days=d.weekday())
    return monday.isoformat()


def week_label(period_id: str) -> str:
    monday = datetime.strptime(period_id, "%Y-%m-%d").date()
    sunday = monday + timedelta(days=6)
    return f"Week of {monday.strftime('%b %d')} – {sunday.strftime('%b %d, %Y')}"


def half_year_key(d: date) -> str:
    half = 1 if d.month <= 6 else 2
    return f"{d.year}-H{half}"


def half_year_label(period_id: str) -> str:
    year_str, half_str = period_id.split("-H")
    months = "Jan – Jun" if half_str == "1" else "Jul – Dec"
    return f"H{half_str} {year_str} ({months})"


def year_key(d: date) -> str:
    return str(d.year)


def year_label(period_id: str) -> str:
    return f"Year {period_id}"


def _render_task_table(story, tasks: list, cell_style) -> None:
    """Shared table-rendering helper used for both per-day and per-period report sections."""
    if not tasks:
        story.append(Paragraph("No tasks recorded.", cell_style))
        return

    header_row = ["Task", "Status", "Created", "Started", "Completed", "Duration"]
    status_display = {"todo": "To Do", "in_progress": "In Progress", "done": "Done"}
    status_colors = {
        "todo": colors.HexColor("#6b6b8a"),
        "in_progress": colors.HexColor("#b8860b"),
        "done": colors.HexColor("#2e9c5c"),
    }

    ordered = sorted(tasks, key=lambda t: (t.get("status", "todo"), t.get("order", 0)))
    table_data = [header_row]
    row_styles = []
    for i, t in enumerate(ordered, start=1):
        table_data.append(
            [
                Paragraph(xml_escape(t.get("text", "")), cell_style),
                status_display.get(t.get("status"), t.get("status", "")),
                fmt_12h(t.get("created_at")),
                fmt_12h(t.get("started_at")),
                fmt_12h(t.get("completed_at")),
                fmt_duration(t.get("created_at"), t.get("completed_at")),
            ]
        )
        row_styles.append(("TEXTCOLOR", (1, i), (1, i), status_colors.get(t.get("status"), colors.black)))

    table = Table(
        table_data,
        colWidths=[1.9 * inch, 0.85 * inch, 0.75 * inch, 0.75 * inch, 0.75 * inch, 0.7 * inch],
        repeatRows=1,
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7c6cff")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dddddd")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5fa")]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                *row_styles,
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 10))


def build_pdf_report(path: str, days: list, tasks_by_day: dict, period_sections: list = None) -> None:
    """Render a PDF report. tasks_by_day maps 'YYYY-MM-DD' -> list of task dicts.
    period_sections (optional) is a list of {"title": str, "tasks": [...]} for the
    weekly / 6-month / yearly boards, appended after the daily sections."""
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle", parent=styles["Heading1"], textColor=colors.HexColor("#4a3aa8")
    )
    date_style = ParagraphStyle(
        "DateHeading", parent=styles["Heading2"], textColor=colors.HexColor("#222222"),
        spaceBefore=14, spaceAfter=6,
    )
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=9, leading=11)

    story = [
        Paragraph("Daily Todo Report", title_style),
        Paragraph(
            f"Generated {datetime.now().strftime('%A, %B %d, %Y')} at {fmt_12h(datetime.now().isoformat())}",
            styles["Normal"],
        ),
        Spacer(1, 14),
    ]

    for day in days:
        try:
            nice_date = datetime.strptime(day, "%Y-%m-%d").strftime("%A, %B %d, %Y")
        except ValueError:
            nice_date = day
        story.append(Paragraph(nice_date, date_style))
        _render_task_table(story, tasks_by_day.get(day, []), cell_style)

    for section in (period_sections or []):
        story.append(Paragraph(section["title"], date_style))
        _render_task_table(story, section.get("tasks", []), cell_style)

    doc = SimpleDocTemplate(path, pagesize=letter, title="Daily Todo Report")
    doc.build(story)


def load_history() -> dict:
    if HISTORY_FILE.exists():
        try:
            return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_history(history: dict) -> None:
    HISTORY_FILE.write_text(json.dumps(history, indent=2), encoding="utf-8")


def archive_day(day: str, tasks: list) -> None:
    """Merge tasks into that day's history entry (doesn't clobber tasks already
    recorded earlier the same day via record_completed_task, e.g. ones since
    deleted or cleared)."""
    history = load_history()
    existing = history.get(day, [])
    existing_ids = {t["id"] for t in existing}
    merged = existing + [t for t in tasks if t["id"] not in existing_ids]
    if merged:
        history[day] = merged
        save_history(history)


def record_completed_task(task: dict) -> None:
    """Permanently record a completed task into today's history the moment it's
    removed from the live board (deleted or cleared), so reports/History still
    show it even though it's gone from today's active list."""
    day = date.today().isoformat()
    history = load_history()
    existing = [t for t in history.get(day, []) if t["id"] != task["id"]]
    existing.append(dict(task))
    history[day] = existing
    save_history(history)


def load_data() -> dict:
    today = date.today().isoformat()
    if TASKS_FILE.exists():
        try:
            data = json.loads(TASKS_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            data = {"date": today, "tasks": [], "next_id": 1}
    else:
        data = {"date": today, "tasks": [], "next_id": 1}

    data.setdefault("next_id", 1)
    data.setdefault("tasks", [])
    data["tasks"] = [migrate_task(t) for t in data["tasks"]]

    if data.get("date") != today:
        archive_day(data["date"], data["tasks"])
        data["tasks"] = [t for t in data["tasks"] if t["status"] != "done"]
        for t in data["tasks"]:
            # scheduled times were for the old day; don't let stale times
            # instantly re-trigger auto-start/reminders on the new day
            t["start_time"] = None
            t["end_time"] = None
            t["reminded_for_end_time"] = None
        data["date"] = today
        save_data(data)

    return data


def save_data(data: dict) -> None:
    TASKS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def load_notes() -> dict:
    if NOTES_FILE.exists():
        try:
            data = json.loads(NOTES_FILE.read_text(encoding="utf-8"))
            data.setdefault("next_id", 1)
            data.setdefault("notes", [])
            data["notes"] = [migrate_note(n) for n in data["notes"]]
            return data
        except (json.JSONDecodeError, OSError):
            pass
    return {"next_id": 1, "notes": []}


def save_notes(data: dict) -> None:
    NOTES_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def load_mail_settings() -> dict:
    if MAIL_WATCHLIST_FILE.exists():
        try:
            data = json.loads(MAIL_WATCHLIST_FILE.read_text(encoding="utf-8"))
            custom_start = data.get("custom_start")
            custom_end = data.get("custom_end")
            return {
                "emails": [e.lower() for e in data.get("emails", [])],
                "hours_window": data.get("hours_window", 2),
                "custom_start": datetime.fromisoformat(custom_start) if custom_start else None,
                "custom_end": datetime.fromisoformat(custom_end) if custom_end else None,
            }
        except (json.JSONDecodeError, OSError, ValueError):
            pass
    return {"emails": [], "hours_window": 2, "custom_start": None, "custom_end": None}


def save_mail_settings(emails: list, hours_window, custom_start=None, custom_end=None) -> None:
    MAIL_WATCHLIST_FILE.write_text(
        json.dumps({
            "emails": emails,
            "hours_window": hours_window,
            "custom_start": custom_start.isoformat() if custom_start else None,
            "custom_end": custom_end.isoformat() if custom_end else None,
        }, indent=2),
        encoding="utf-8",
    )


def load_portals() -> list:
    if PORTALS_FILE.exists():
        try:
            data = json.loads(PORTALS_FILE.read_text(encoding="utf-8"))
            return data.get("portals", [])
        except (json.JSONDecodeError, OSError):
            pass
    return []


def save_portals(portals: list) -> None:
    PORTALS_FILE.write_text(json.dumps({"portals": portals}, indent=2), encoding="utf-8")


def make_tray_image() -> Image.Image:
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((2, 2, size - 2, size - 2), fill=(124, 108, 255, 255))
    d.line((18, 33, 28, 44), fill=(255, 255, 255, 255), width=6)
    d.line((28, 44, 47, 20), fill=(255, 255, 255, 255), width=6)
    return img


class PeriodTaskBoard:
    """A Kanban-style task board scoped to a repeating period (week / half-year / year).

    Mirrors the daily board's card/drag/priority-grouping UI, minus the daily
    board's time-of-day scheduling & reminders (those don't apply to a task
    spanning weeks or months). Each instance owns its own data + history file
    and rolls over automatically when the period changes.
    """

    def __init__(self, app, add_placeholder, data_file: Path, history_file: Path,
                 period_key_fn, period_label_fn, strike_font, normal_font):
        self.app = app
        self.add_placeholder = add_placeholder
        self.data_file = data_file
        self.history_file = history_file
        self.period_key_fn = period_key_fn
        self.period_label_fn = period_label_fn
        self.strike_font = strike_font
        self.normal_font = normal_font

        self.card_widgets = {}
        self.column_frames = {}
        self.column_containers = {}
        self.column_headers = {}
        self.drag_data = {"task_id": None, "start_x": 0, "start_y": 0, "ghost": None, "dragging": False}

        self.data = self._load_data()

    # ---- persistence -----------------------------------------------
    def _load_history(self) -> dict:
        if self.history_file.exists():
            try:
                return json.loads(self.history_file.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                return {}
        return {}

    def _save_history(self, history: dict) -> None:
        self.history_file.write_text(json.dumps(history, indent=2), encoding="utf-8")

    def _archive_period(self, period_id: str, tasks: list) -> None:
        history = self._load_history()
        existing = history.get(period_id, [])
        existing_ids = {t["id"] for t in existing}
        merged = existing + [t for t in tasks if t["id"] not in existing_ids]
        if merged:
            history[period_id] = merged
            self._save_history(history)

    def _record_completed(self, task: dict) -> None:
        period_id = self.period_key_fn(date.today())
        history = self._load_history()
        existing = [t for t in history.get(period_id, []) if t["id"] != task["id"]]
        existing.append(dict(task))
        history[period_id] = existing
        self._save_history(history)

    def _load_data(self) -> dict:
        current_period = self.period_key_fn(date.today())
        if self.data_file.exists():
            try:
                data = json.loads(self.data_file.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                data = {"period": current_period, "tasks": [], "next_id": 1}
        else:
            data = {"period": current_period, "tasks": [], "next_id": 1}

        data.setdefault("next_id", 1)
        data.setdefault("tasks", [])
        data["tasks"] = [migrate_task(t) for t in data["tasks"]]

        if data.get("period") != current_period:
            self._archive_period(data["period"], data["tasks"])
            data["tasks"] = [t for t in data["tasks"] if t["status"] != "done"]
            data["period"] = current_period
            self._save_data(data)

        return data

    def _save_data(self, data: dict = None) -> None:
        data = data if data is not None else self.data
        self.data_file.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def reload(self) -> None:
        self.data = self._load_data()

    def get_report_tasks(self) -> list:
        """Current period's tasks: live board plus anything already completed
        and removed this period (so a report never loses them)."""
        current_period = self.period_key_fn(date.today())
        history = self._load_history()
        live = self.data["tasks"]
        live_ids = {t["id"] for t in live}
        already_recorded = [t for t in history.get(current_period, []) if t["id"] not in live_ids]
        return [migrate_task(dict(t)) for t in live + already_recorded]

    def report_section(self) -> dict:
        current_period = self.period_key_fn(date.today())
        return {"title": self.period_label_fn(current_period), "tasks": self.get_report_tasks()}

    # ---- UI construction ---------------------------------------------
    def build_tab(self, root) -> None:
        header = ctk.CTkFrame(root, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(22, 8))

        self.period_label = ctk.CTkLabel(
            header, text="", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY
        )
        self.period_label.pack(anchor="w")

        self.subtitle_label = ctk.CTkLabel(
            header, text="", font=("Segoe UI", 13), text_color=TEXT_SECONDARY
        )
        self.subtitle_label.pack(anchor="w", pady=(2, 0))

        progress_frame = ctk.CTkFrame(root, fg_color="transparent")
        progress_frame.pack(fill="x", padx=24, pady=(2, 4))
        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, height=10, corner_radius=6, fg_color=CARD, progress_color=ACCENT
        )
        self.progress_bar.pack(fill="x", side="left", expand=True)
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(
            progress_frame, text="0%", font=("Segoe UI", 12, "bold"), text_color=ACCENT, width=40
        )
        self.progress_label.pack(side="left", padx=(10, 0))

        entry_frame = ctk.CTkFrame(root, fg_color="transparent")
        entry_frame.pack(fill="x", padx=24, pady=(16, 10))
        self.entry = ctk.CTkEntry(
            entry_frame, placeholder_text=self.add_placeholder, fg_color=ENTRY_BG,
            border_color=BORDER, border_width=1, corner_radius=12, height=42, font=("Segoe UI", 13),
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry.bind("<Return>", lambda _e: self.add_task())

        add_btn = self.app._floating_circle_button(
            entry_frame, "+", 44, ACCENT, ACCENT_HOVER, "white", ("Segoe UI", 20, "bold"), self.add_task
        )
        add_btn.pack(side="left")

        board = ctk.CTkFrame(root, fg_color="transparent")
        board.pack(fill="both", expand=True, padx=18, pady=(4, 4))
        board.grid_columnconfigure(0, weight=1, uniform="col")
        board.grid_columnconfigure(1, weight=1, uniform="col")
        board.grid_columnconfigure(2, weight=1, uniform="col")
        board.grid_rowconfigure(0, weight=1)

        for i, status in enumerate(STATUS_ORDER):
            container = ctk.CTkFrame(board, fg_color=CARD, corner_radius=14, border_width=0, border_color=ACCENT)
            container.grid(row=0, column=i, sticky="nsew", padx=6)
            container.grid_rowconfigure(1, weight=1)
            container.grid_columnconfigure(0, weight=1)
            self.column_containers[status] = container

            col_header = ctk.CTkFrame(container, fg_color="transparent")
            col_header.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))
            dot = ctk.CTkFrame(col_header, width=9, height=9, corner_radius=5, fg_color=COLUMN_ACCENTS[status])
            dot.pack(side="left", padx=(0, 8))
            dot.pack_propagate(False)
            ctk.CTkLabel(
                col_header, text=STATUS_LABELS[status], font=("Segoe UI", 13, "bold"), text_color=TEXT_PRIMARY,
            ).pack(side="left")
            count_lbl = ctk.CTkLabel(col_header, text="0", font=("Segoe UI", 12), text_color=TEXT_SECONDARY)
            count_lbl.pack(side="right")
            self.column_headers[status] = count_lbl

            scroll = ctk.CTkScrollableFrame(container, fg_color="transparent", scrollbar_button_color=BORDER)
            scroll.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 12))
            self.column_frames[status] = scroll

        footer = ctk.CTkFrame(root, fg_color="transparent")
        footer.pack(fill="x", padx=24, pady=(4, 18))
        self.count_label = ctk.CTkLabel(footer, text="", font=("Segoe UI", 12), text_color=TEXT_SECONDARY)
        self.count_label.pack(side="left")
        ctk.CTkButton(
            footer, text="Clear completed", fg_color="transparent", hover_color=CARD,
            text_color=TEXT_SECONDARY, font=("Segoe UI", 12), width=120, height=26, corner_radius=8,
            command=self.clear_completed,
        ).pack(side="right")
        ctk.CTkLabel(
            footer, text="Drag a card between columns · click a card for details & notes",
            font=("Segoe UI", 11), text_color=TEXT_MUTED,
        ).pack(side="right", padx=(0, 16))

        self.render()

    # ---- data operations ------------------------------------------------
    def _find_task(self, task_id):
        for t in self.data["tasks"]:
            if t["id"] == task_id:
                return t
        return None

    def add_task(self):
        text = self.entry.get().strip()
        if not text:
            return
        todo_orders = [t.get("order", 0) for t in self.data["tasks"] if t["status"] == "todo"]
        task = {
            "id": self.data["next_id"], "text": text, "notes": "", "status": "todo",
            "priority": "medium", "order": (max(todo_orders) + 1) if todo_orders else 0,
            "created_at": datetime.now().isoformat(), "started_at": None, "completed_at": None,
            "start_time": None, "end_time": None, "reminded_for_end_time": None,
        }
        self.data["next_id"] += 1
        self.data["tasks"].append(task)
        self.entry.delete(0, "end")
        self._save_data()
        self.render()

    def delete_task(self, task_id: int):
        task = self._find_task(task_id)
        if task is not None and task["status"] == "done":
            self._record_completed(task)
        self.data["tasks"] = [t for t in self.data["tasks"] if t["id"] != task_id]
        self._save_data()
        self.render()

    def clear_completed(self):
        for t in self.data["tasks"]:
            if t["status"] == "done":
                self._record_completed(t)
        self.data["tasks"] = [t for t in self.data["tasks"] if t["status"] != "done"]
        self._save_data()
        self.render()

    # ---- rendering --------------------------------------------------------
    def render(self):
        period_id = self.period_key_fn(date.today())
        self.period_label.configure(text=self.period_label_fn(period_id))

        tasks = self.data["tasks"]
        total = len(tasks)
        done = sum(1 for t in tasks if t["status"] == "done")
        self.subtitle_label.configure(
            text="All clear ✨" if total == 0 else f"{total - done} task(s) remaining"
        )
        self.progress_bar.set(done / total if total else 0)
        self.progress_label.configure(text=f"{round(done / total * 100) if total else 0}%")
        self.count_label.configure(text=f"{done} of {total} completed")

        self.card_widgets = {}
        for status in STATUS_ORDER:
            frame = self.column_frames[status]
            for child in frame.winfo_children():
                child.destroy()

            subset = [t for t in tasks if t["status"] == status]
            self.column_headers[status].configure(text=str(len(subset)))

            if not subset:
                ctk.CTkLabel(
                    frame, text="Drop tasks here", font=("Segoe UI", 11), text_color=TEXT_MUTED
                ).pack(pady=18)
                continue

            for priority in PRIORITY_ORDER:
                group = sorted(
                    [t for t in subset if t.get("priority", "medium") == priority],
                    key=lambda t: t.get("order", 0),
                )
                if not group:
                    continue
                group_header = ctk.CTkFrame(frame, fg_color="transparent")
                group_header.pack(fill="x", pady=(10, 2))
                ctk.CTkLabel(
                    group_header, text="", width=8, height=8, corner_radius=4,
                    fg_color=PRIORITY_COLORS[priority],
                ).pack(side="left", padx=(2, 6))
                ctk.CTkLabel(
                    group_header, text=f"{PRIORITY_LABELS[priority]} priority",
                    font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED,
                ).pack(side="left")
                for t in group:
                    self.card_widgets[t["id"]] = self._build_card(frame, t)

    def _build_card(self, parent, task: dict):
        status = task["status"]
        card = ctk.CTkFrame(parent, fg_color=CARD_DONE if status == "done" else CARD2, corner_radius=10)
        card.pack(fill="x", pady=5, padx=2)

        stripe = ctk.CTkFrame(card, width=4, corner_radius=0, fg_color=COLUMN_ACCENTS[status])
        stripe.pack(side="left", fill="y")
        stripe.pack_propagate(False)

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True, padx=(10, 4), pady=8)

        priority = task.get("priority", "medium")
        badge = ctk.CTkLabel(
            content, text=PRIORITY_LABELS[priority], font=("Segoe UI", 9, "bold"),
            text_color=PRIORITY_TEXT_COLORS[priority], fg_color=PRIORITY_COLORS[priority], corner_radius=6,
        )
        badge.pack(anchor="w", pady=(0, 5), ipadx=6, ipady=1)

        title_lbl = ctk.CTkLabel(
            content, text=task["text"], font=self.strike_font if status == "done" else self.normal_font,
            text_color=TEXT_MUTED if status == "done" else TEXT_PRIMARY, anchor="w", justify="left", wraplength=185,
        )
        title_lbl.pack(fill="x", anchor="w")

        drag_targets = [card, content, badge, title_lbl]

        if task.get("notes"):
            preview = task["notes"].splitlines()[0][:70]
            note_lbl = ctk.CTkLabel(
                content, text=f"\U0001F4DD {preview}", font=("Segoe UI", 10), text_color=TEXT_SECONDARY,
                anchor="w", justify="left", wraplength=185,
            )
            note_lbl.pack(fill="x", anchor="w", pady=(4, 0))
            drag_targets.append(note_lbl)

        del_btn = self.app._floating_circle_button(
            card, "✕", 22, CARD2, DANGER, TEXT_SECONDARY, ("Segoe UI", 11),
            lambda tid=task["id"]: self.delete_task(tid),
        )
        del_btn.pack(side="right", padx=(0, 8), pady=8, anchor="n")

        for widget in drag_targets:
            self._bind_drag(widget, task["id"])
            widget.configure(cursor="hand2")

        return card

    # ---- drag and drop -----------------------------------------------
    def _bind_drag(self, widget, task_id: int):
        widget.bind("<ButtonPress-1>", lambda e, tid=task_id: self._on_drag_start(e, tid))
        widget.bind("<B1-Motion>", self._on_drag_motion)
        widget.bind("<ButtonRelease-1>", self._on_drag_release)

    def _on_drag_start(self, event, task_id: int):
        self.drag_data = {
            "task_id": task_id, "start_x": event.x_root, "start_y": event.y_root,
            "ghost": None, "dragging": False,
        }

    def _on_drag_motion(self, event):
        if self.drag_data["task_id"] is None:
            return
        dx = abs(event.x_root - self.drag_data["start_x"])
        dy = abs(event.y_root - self.drag_data["start_y"])
        if not self.drag_data["dragging"] and (dx > 6 or dy > 6):
            self.drag_data["dragging"] = True
            self._create_ghost()
        if self.drag_data["dragging"] and self.drag_data["ghost"]:
            self.drag_data["ghost"].geometry(f"+{event.x_root + 14}+{event.y_root + 12}")
            self._highlight_drop_target(event)

    def _create_ghost(self):
        task = self._find_task(self.drag_data["task_id"])
        if task is None:
            return
        ghost = ctk.CTkToplevel(self.app.root)
        ghost.overrideredirect(True)
        ghost.attributes("-topmost", True)
        try:
            ghost.attributes("-alpha", 0.88)
        except Exception:
            pass
        ctk.CTkLabel(
            ghost, text=task["text"], fg_color=ACCENT, text_color="white", corner_radius=10,
            padx=14, pady=10, font=("Segoe UI", 12, "bold"),
        ).pack()
        self.drag_data["ghost"] = ghost

    def _column_at(self, x_root: int, y_root: int):
        for status, container in self.column_containers.items():
            fx, fy = container.winfo_rootx(), container.winfo_rooty()
            fw, fh = container.winfo_width(), container.winfo_height()
            if fx <= x_root <= fx + fw and fy <= y_root <= fy + fh:
                return status
        return None

    def _highlight_drop_target(self, event):
        target = self._column_at(event.x_root, event.y_root)
        for status, container in self.column_containers.items():
            container.configure(border_width=2 if status == target else 0)

    def _on_drag_release(self, event):
        task_id = self.drag_data.get("task_id")
        dragging = self.drag_data.get("dragging")
        ghost = self.drag_data.get("ghost")
        if ghost:
            ghost.destroy()
        for container in self.column_containers.values():
            container.configure(border_width=0)

        if task_id is None:
            self.drag_data = {"task_id": None, "start_x": 0, "start_y": 0, "ghost": None, "dragging": False}
            return

        if dragging:
            target_status = self._column_at(event.x_root, event.y_root)
            if target_status:
                self._move_task(task_id, target_status, event.y_root)
        else:
            self.open_task_dialog(task_id)

        self.drag_data = {"task_id": None, "start_x": 0, "start_y": 0, "ghost": None, "dragging": False}

    def _move_task(self, task_id: int, target_status: str, drop_y_root: int = 10 ** 9):
        task = self._find_task(task_id)
        if task is None:
            return
        old_status = task["status"]

        siblings = [t for t in self.data["tasks"] if t["status"] == target_status and t["id"] != task_id]
        positions = []
        for s in siblings:
            w = self.card_widgets.get(s["id"])
            if w is not None and w.winfo_exists():
                cy = w.winfo_rooty() + w.winfo_height() / 2
                positions.append((cy, s))
        positions.sort(key=lambda p: p[0])

        idx = 0
        for cy, _s in positions:
            if drop_y_root > cy:
                idx += 1
            else:
                break

        if not positions:
            new_order = 0
        elif idx == 0:
            new_order = positions[0][1].get("order", 0) - 1
        elif idx >= len(positions):
            new_order = positions[-1][1].get("order", 0) + 1
        else:
            a = positions[idx - 1][1].get("order", 0)
            b = positions[idx][1].get("order", 0)
            new_order = (a + b) / 2

        now_iso = datetime.now().isoformat()
        if target_status == "in_progress" and not task.get("started_at"):
            task["started_at"] = now_iso
        if target_status == "done":
            task["completed_at"] = now_iso
        elif old_status == "done":
            task["completed_at"] = None

        task["status"] = target_status
        task["order"] = new_order

        self._save_data()
        self.render()

    # ---- task detail dialog --------------------------------------------
    def open_task_dialog(self, task_id: int):
        task = self._find_task(task_id)
        if task is None:
            return

        win = ctk.CTkToplevel(self.app.root)
        win.title("Task Details")
        win.geometry("380x480")
        win.configure(fg_color=BG)
        win.transient(self.app.root)
        win.grab_set()
        win.resizable(False, False)

        ctk.CTkLabel(
            win, text="Title", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(20, 4))
        title_entry = ctk.CTkEntry(
            win, fg_color=ENTRY_BG, border_color=BORDER, border_width=1, corner_radius=10, height=38
        )
        title_entry.insert(0, task["text"])
        title_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(
            win, text="Notes", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(16, 4))
        notes_box = ctk.CTkTextbox(
            win, fg_color=ENTRY_BG, border_color=BORDER, border_width=1, corner_radius=10,
            height=170, font=("Segoe UI", 12),
        )
        notes_box.insert("1.0", task.get("notes", ""))
        notes_box.pack(fill="both", padx=20, expand=True)

        ctk.CTkLabel(
            win, text="Priority", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(16, 4))
        priority_row = ctk.CTkFrame(win, fg_color="transparent")
        priority_row.pack(fill="x", padx=20, pady=(0, 10))
        current_priority = task.get("priority", "medium")
        priority_dot = ctk.CTkLabel(
            priority_row, text="", width=14, height=14, corner_radius=7,
            fg_color=PRIORITY_COLORS[current_priority],
        )
        priority_dot.pack(side="left", padx=(0, 8))
        label_to_priority = {v: k for k, v in PRIORITY_LABELS.items()}

        def on_priority_change(choice):
            priority_dot.configure(fg_color=PRIORITY_COLORS[label_to_priority[choice]])

        priority_menu = ctk.CTkOptionMenu(
            priority_row, values=[PRIORITY_LABELS[p] for p in PRIORITY_ORDER], fg_color=ENTRY_BG,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER, corner_radius=10, command=on_priority_change,
        )
        priority_menu.set(PRIORITY_LABELS[current_priority])
        priority_menu.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            win, text="Status", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(0, 4))
        status_menu = ctk.CTkOptionMenu(
            win, values=[STATUS_LABELS[s] for s in STATUS_ORDER], fg_color=ENTRY_BG,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER, corner_radius=10,
        )
        status_menu.set(STATUS_LABELS[task["status"]])
        status_menu.pack(fill="x", padx=20, pady=(0, 10))

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(10, 20), side="bottom")

        def do_delete():
            win.destroy()
            self.delete_task(task_id)

        def do_save():
            new_text = title_entry.get().strip()
            if not new_text:
                return
            task["text"] = new_text
            task["notes"] = notes_box.get("1.0", "end-1c").strip()
            task["priority"] = label_to_priority[priority_menu.get()]

            label_to_status = {v: k for k, v in STATUS_LABELS.items()}
            new_status = label_to_status[status_menu.get()]
            if new_status != task["status"]:
                now_iso = datetime.now().isoformat()
                if new_status == "in_progress" and not task.get("started_at"):
                    task["started_at"] = now_iso
                if new_status == "done":
                    task["completed_at"] = now_iso
                elif task["status"] == "done":
                    task["completed_at"] = None
                existing = [
                    t.get("order", 0) for t in self.data["tasks"]
                    if t["status"] == new_status and t["id"] != task_id
                ]
                task["order"] = (max(existing) + 1) if existing else 0
                task["status"] = new_status

            self._save_data()
            win.destroy()
            self.render()

        ctk.CTkButton(
            btn_row, text="Delete", fg_color="transparent", hover_color=DANGER, text_color=DANGER_TEXT,
            corner_radius=8, width=70, command=do_delete,
        ).pack(side="left")
        ctk.CTkButton(
            btn_row, text="Cancel", fg_color="transparent", hover_color=CARD, text_color=TEXT_SECONDARY,
            corner_radius=8, width=70, command=win.destroy,
        ).pack(side="right", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Save", fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8,
            width=80, command=do_save,
        ).pack(side="right")


class MailMonitor:
    """Read-only Outlook inbox monitor via local COM automation.

    Uses whichever Outlook desktop profile is already signed in on this
    machine — this app never sees or stores a password. Clicking a message
    opens it in the real Outlook window; nothing is ever sent from here.
    """

    MAX_ITEMS = 30
    SCAN_LIMIT = 500  # hard safety cap; the time-window cutoff below usually stops far earlier
    REFRESH_MS = 120_000  # 2 minutes

    HOURS_OPTIONS = {
        "All": None,
        "Last 1 hour": 1, "Last 2 hours": 2, "Last 6 hours": 6, "Last 12 hours": 12,
        "Last 24 hours": 24, "Last 3 days": 72, "Last 7 days": 168,
        "Custom range...": "custom",
    }
    CUSTOM_DATE_RANGE_DAYS = 90  # how far back the date dropdowns go

    RULES_TAB_NAME = "⚙ Rules"

    def __init__(self, app):
        self.app = app
        self._namespace = None
        self.session_start = datetime.now()
        self.notified_ids: set = set()
        self.status_label = None
        self.account_label = None
        self.chips_frame = None
        self.watch_entry = None
        self.hours_menu = None
        self.custom_range_frame = None
        self.custom_start_date_menu = None
        self.custom_start_hour_menu = None
        self.custom_start_min_menu = None
        self.custom_end_date_menu = None
        self.custom_end_hour_menu = None
        self.custom_end_min_menu = None
        self._date_options: dict = {}  # display label -> date
        self.emails: list = []

        self.mail_tabs_container = None
        self.rule_tabview = None
        self.rule_frames: dict = {}       # email -> its scrollable frame
        self.rule_tab_labels: dict = {}   # email -> tab label used (for lookups)
        self._built_rule_keys: list = []  # watchlist snapshot the current tabs were built from

        self.rule_folders: dict = {}      # folder name -> Outlook folder COM object (from mail rules)
        self.folder_frames: dict = {}     # folder name -> its scrollable frame
        self.folder_mail: dict = {}       # folder name -> list of email dicts currently shown there
        self._built_folder_keys: list = []

        settings = load_mail_settings()
        self.watchlist: list = settings["emails"]
        self.hours_window = settings["hours_window"]
        self.custom_start = settings["custom_start"]
        self.custom_end = settings["custom_end"]

    def _get_namespace(self):
        if self._namespace is None:
            outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = outlook.GetNamespace("MAPI")
        return self._namespace

    def _sender_email(self, item) -> str:
        """Resolve the sender's SMTP address, following Exchange DN addresses
        (internal senders often report as an X.500 DN, not a plain SMTP address)."""
        addr = (getattr(item, "SenderEmailAddress", "") or "").strip()
        if addr and not addr.startswith("/"):
            return addr
        try:
            exch_user = item.Sender.GetExchangeUser()
            if exch_user is not None and exch_user.PrimarySmtpAddress:
                return exch_user.PrimarySmtpAddress
        except Exception:
            pass
        return addr

    EXCHANGE_DL_TYPE = 1  # olExchangeDistributionListAddressEntry

    def _lookup_group(self, ns, name: str):
        """Look up a Distribution List / group by exact display name in the
        Global Address List. NOTE: AddressEntries.Item(name) does a fuzzy
        match and will silently return an unrelated entry for a name that
        doesn't exist (confirmed empirically) — so the returned entry's own
        .Name must be checked to exactly match what was requested."""
        try:
            entry = ns.GetGlobalAddressList().AddressEntries.Item(name)
            if entry.Name.strip().lower() == name.strip().lower():
                return entry
        except Exception:
            pass
        return None

    def _build_match_map(self, ns) -> dict:
        """Map every raw address form back to the original watchlist entry it
        came from — resolved once per refresh, not per scanned message — so a
        matched message can be filed under the correct 'rule' tab.

        A watchlist entry can be a plain email OR a Distribution List / group
        name (e.g. "US ADMX GDL All") — in that case every member's address is
        mapped back to the group's own entry, so watching a group covers mail
        to/from anyone in it. Members' raw .Address (often an Exchange DN, not
        SMTP) is used directly rather than resolved via GetExchangeUser() per
        member — that DN is a stable identifier, so it matches exactly what a
        message's own Sender/Recipients would show for that same person,
        without needing a directory lookup per member (which would be slow
        for a large group).

        CreateRecipient().Resolve() reliably resolves plain email addresses to
        their Exchange DN, but does NOT reliably resolve a group by display
        name (confirmed empirically) — group names go through the GAL lookup
        above instead."""
        match_map = {}
        for entry in self.watchlist:
            match_map[entry.lower()] = entry
            if "@" in entry:
                try:
                    recip = ns.CreateRecipient(entry)
                    recip.Resolve()
                    if recip.Resolved and recip.AddressEntry and recip.AddressEntry.Address:
                        match_map[recip.AddressEntry.Address.lower()] = entry
                except Exception:
                    pass
                continue

            ae = self._lookup_group(ns, entry)
            if ae is None:
                continue
            if ae.Address:
                match_map[ae.Address.lower()] = entry
            if ae.AddressEntryUserType == self.EXCHANGE_DL_TYPE:
                try:
                    members = ae.Members
                    for i in range(1, members.Count + 1):
                        member_addr = (members.Item(i).Address or "").lower()
                        if member_addr:
                            match_map[member_addr] = entry
                except Exception:
                    pass
        return match_map

    # olFolderInbox=6 (ReceivedTime is meaningful); olFolderSentMail=5 (use SentOn —
    # a sent item's ReceivedTime is meaningless, so watching an address only ever
    # caught mail it sent you, never mail you sent it, until this was added).
    FOLDERS = [(6, "ReceivedTime", "inbox"), (5, "SentOn", "sent")]

    def _discover_rule_folders(self, ns) -> dict:
        """Find folders that an Outlook mail rule files messages into directly
        (e.g. a rule moving mail sent to a Distribution List into a folder
        named "EH") — that mail never reaches Inbox/Sent Items at all, so the
        watchlist-based scan above can never see it no matter what. These
        become their own dedicated tabs showing that folder's contents
        directly; Outlook's own rule is already the filter, so no watchlist
        matching is needed for them."""
        folders = {}
        try:
            rules = ns.DefaultStore.GetRules()
        except Exception:
            return folders
        for i in range(1, rules.Count + 1):
            try:
                rule = rules.Item(i)
                if not rule.Enabled:
                    continue
                move_action = rule.Actions.MoveToFolder
                if move_action.Enabled:
                    folder = move_action.Folder
                    folders[folder.Name] = folder
            except Exception:
                continue
        return folders

    def _compute_cutoffs(self):
        """Returns (start_cutoff, end_cutoff) — either may be None (no bound
        on that side). Presets only bound the start (end = now, implicitly);
        a custom range bounds both sides explicitly."""
        if self.hours_window == "custom":
            return self.custom_start, self.custom_end
        if self.hours_window is None:
            return None, None
        return datetime.now() - timedelta(hours=self.hours_window), None

    def _fetch_folder_recent(self, folder, start_cutoff, end_cutoff) -> list:
        """Like _scan_folder, but for a rule-target folder: no watchlist
        matching needed (the Outlook rule that files mail here already is
        the filter) — just return everything in the window, bounded the
        same way (SCAN_LIMIT, time cutoff)."""
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        results = []
        scanned = 0
        for item in items:
            if scanned >= self.SCAN_LIMIT or len(results) >= self.MAX_ITEMS:
                break
            scanned += 1
            try:
                if item.Class != 43:  # olMail
                    continue
                rt = item.ReceivedTime
                item_dt = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
                if end_cutoff is not None and item_dt > end_cutoff:
                    continue  # too new — skip, but keep scanning for older matches
                if start_cutoff is not None and item_dt < start_cutoff:
                    break  # too old, and sorted newest-first, so nothing further can be in range
                results.append(self._build_result_dict(item, item_dt, "rule_folder", set()))
            except Exception:
                continue
        return results

    def _build_result_dict(self, item, item_dt, folder_label: str, matched_rules: set) -> dict:
        sender_email = self._sender_email(item)
        try:
            attachment_count = item.Attachments.Count
        except Exception:
            attachment_count = 0
        try:
            importance_high = item.Importance == 2  # olImportanceHigh
        except Exception:
            importance_high = False
        try:
            to_names = [rcp.Name for rcp in item.Recipients if rcp.Type == 1]  # olTo
        except Exception:
            to_names = []
        preview = " ".join((item.Body or "").split())[:220]
        return {
            "id": item.EntryID,
            "sender": item.SenderName or sender_email or "(unknown sender)",
            "sender_email": sender_email,
            "subject": item.Subject or "(no subject)",
            "received": item_dt,
            "folder": folder_label,
            "unread": bool(item.UnRead),
            "preview": preview,
            "attachment_count": attachment_count,
            "importance_high": importance_high,
            "to_names": to_names,
            "matched_rules": sorted(matched_rules),
        }

    def _scan_folder(self, ns, folder_id: int, time_attr: str, folder_label: str,
                      match_map: dict, start_cutoff, end_cutoff) -> list:
        """Bounded scan (SCAN_LIMIT items, further bounded by the time cutoff)
        checking BOTH sender and recipient matches. Sender matches are also
        covered — without the SCAN_LIMIT bound — by _restrict_sender_matches
        below; this scan is what actually finds recipient (To/CC) matches,
        since Outlook's Restrict can only filter To/CC by display name, not
        by address (confirmed empirically), so there's no fast/unbounded
        equivalent for that half of the matching.

        Deliberately does NOT stop early once some number of matches is
        found: a broadly-matching rule (e.g. watching your own address, which
        matches nearly every inbox message) would otherwise consume the whole
        budget before the scan ever reaches a less-common rule's matches —
        capping happens per-rule afterward in _fetch_recent, not here."""
        folder = ns.GetDefaultFolder(folder_id)
        items = folder.Items
        items.Sort(f"[{time_attr}]", True)

        results = []
        scanned = 0
        for item in items:
            if scanned >= self.SCAN_LIMIT:
                break
            scanned += 1
            try:
                if item.Class != 43:  # olMail
                    continue
                ts = getattr(item, time_attr)
                item_dt = datetime(ts.year, ts.month, ts.day, ts.hour, ts.minute, ts.second)
                if end_cutoff is not None and item_dt > end_cutoff:
                    continue  # too new — skip, but keep scanning for older matches
                if start_cutoff is not None and item_dt < start_cutoff:
                    # sorted newest-first, so nothing after this can be in the window either
                    break
                # cheap raw-string matching first (no per-item Exchange resolution)
                sender_raw = (getattr(item, "SenderEmailAddress", "") or "").lower()
                matched_rules = set()
                if sender_raw in match_map:
                    matched_rules.add(match_map[sender_raw])
                try:
                    recipient_raws = {(rcp.Address or "").lower() for rcp in item.Recipients}
                except Exception:
                    recipient_raws = set()
                for raw in recipient_raws:
                    if raw in match_map:
                        matched_rules.add(match_map[raw])
                if not matched_rules:
                    continue
                results.append(self._build_result_dict(item, item_dt, folder_label, matched_rules))
            except Exception:
                continue
        return results

    def _restrict_sender_matches(self, ns, folder_id: int, time_attr: str, folder_label: str,
                                  match_map: dict, start_cutoff, end_cutoff) -> list:
        """Fast, server-side search for SENDER matches across the WHOLE
        folder — not bounded by SCAN_LIMIT — since Outlook's Restrict works
        reliably for SenderEmailAddress equality (verified against a real
        526-member group: 2.6s for the whole inbox vs. 36s to linearly scan
        just 2000 items). This is what lets a rarely-matching rule (e.g. a
        notification account that only ever appears as a sender) still be
        found even when it's far older than SCAN_LIMIT would otherwise reach."""
        if not match_map:
            return []
        folder = ns.GetDefaultFolder(folder_id)
        clauses = [f"[SenderEmailAddress] = '{addr.replace(chr(39), chr(39) * 2)}'" for addr in match_map]
        try:
            restricted = folder.Items.Restrict(" OR ".join(clauses))
        except Exception:
            return []

        results = []
        for item in restricted:
            try:
                if item.Class != 43:
                    continue
                sender_raw = (getattr(item, "SenderEmailAddress", "") or "").lower()
                rule = match_map.get(sender_raw)
                if rule is None:
                    continue
                ts = getattr(item, time_attr)
                item_dt = datetime(ts.year, ts.month, ts.day, ts.hour, ts.minute, ts.second)
                if end_cutoff is not None and item_dt > end_cutoff:
                    continue
                if start_cutoff is not None and item_dt < start_cutoff:
                    continue
                results.append(self._build_result_dict(item, item_dt, folder_label, {rule}))
            except Exception:
                continue
        return results

    def _fetch_recent(self):
        ns = self._get_namespace()
        start_cutoff, end_cutoff = self._compute_cutoffs()
        match_map = self._build_match_map(ns)

        results = []
        for folder_id, time_attr, folder_label in self.FOLDERS:
            results.extend(self._restrict_sender_matches(
                ns, folder_id, time_attr, folder_label, match_map, start_cutoff, end_cutoff
            ))
            results.extend(self._scan_folder(
                ns, folder_id, time_attr, folder_label, match_map, start_cutoff, end_cutoff
            ))

        # the same message can be found by both passes above — merge, don't duplicate
        by_id = {}
        for e in results:
            existing = by_id.get(e["id"])
            if existing is None:
                by_id[e["id"]] = e
            else:
                existing["matched_rules"] = sorted(set(existing["matched_rules"]) | set(e["matched_rules"]))
        merged = sorted(by_id.values(), key=lambda e: e["received"], reverse=True)

        # Cap per rule, not globally — otherwise a broadly-matching rule (e.g.
        # your own address) starves every other rule's tab of a fair share.
        per_rule_counts: dict = {}
        final = []
        for e in merged:
            eligible = [r for r in e["matched_rules"] if per_rule_counts.get(r, 0) < self.MAX_ITEMS]
            if not eligible:
                continue
            for r in eligible:
                per_rule_counts[r] = per_rule_counts.get(r, 0) + 1
            final.append(e)
        return final, ns.CurrentUser

    def refresh(self):
        if not OUTLOOK_AVAILABLE:
            return

        current_user = None
        try:
            ns = self._get_namespace()
            current_user = ns.CurrentUser
            start_cutoff, end_cutoff = self._compute_cutoffs()
            self.rule_folders = self._discover_rule_folders(ns)
            self.folder_mail = {
                name: self._fetch_folder_recent(folder, start_cutoff, end_cutoff)
                for name, folder in self.rule_folders.items()
            }
        except Exception as e:
            self._show_error(str(e))
            self.app.root.after(self.REFRESH_MS, self.refresh)
            return

        if not self.watchlist:
            self.emails = []
        else:
            try:
                self.emails, current_user = self._fetch_recent()
            except Exception as e:
                self._show_error(str(e))
                self.app.root.after(self.REFRESH_MS, self.refresh)
                return
            self._notify_new_mail(self.emails)

        self._render(current_user)
        self.app.root.after(self.REFRESH_MS, self.refresh)

    # ---- watchlist / time-window management ----------------------------
    def add_watch_email(self):
        raw = self.watch_entry.get().strip()
        if not raw:
            return

        if "@" in raw:
            entry = raw.lower()
        else:
            # not a plain email — try resolving it as a Distribution List /
            # group name (e.g. "US ADMX GDL All") via the Global Address List
            entry = None
            if OUTLOOK_AVAILABLE:
                ae = self._lookup_group(self._get_namespace(), raw)
                if ae is not None and ae.AddressEntryUserType == self.EXCHANGE_DL_TYPE:
                    entry = ae.Name
            if entry is None:
                return  # not an email and not a resolvable group — ignore

        if entry not in self.watchlist:
            self.watchlist.append(entry)
            save_mail_settings(self.watchlist, self.hours_window)
        self.watch_entry.delete(0, "end")
        self._render_chips()
        self.refresh()

    def remove_watch_email(self, email: str):
        self.watchlist = [e for e in self.watchlist if e != email]
        save_mail_settings(self.watchlist, self.hours_window)
        self._render_chips()
        self.refresh()

    def set_hours_window(self, choice: str):
        self.hours_window = self.HOURS_OPTIONS[choice]
        if self.custom_range_frame is not None:
            if self.hours_window == "custom":
                self.custom_range_frame.pack(fill="x", padx=24, pady=(0, 6))
                return  # wait for "Apply" — don't refresh with a stale/no range yet
            self.custom_range_frame.pack_forget()
        save_mail_settings(self.watchlist, self.hours_window, self.custom_start, self.custom_end)
        self.refresh()

    def _build_date_options(self) -> dict:
        today = date.today()
        options = {}
        for i in range(self.CUSTOM_DATE_RANGE_DAYS + 1):
            d = today - timedelta(days=i)
            options[d.strftime("%b %d, %Y")] = d
        return options

    def apply_custom_range(self) -> None:
        try:
            start_date = self._date_options[self.custom_start_date_menu.get()]
            end_date = self._date_options[self.custom_end_date_menu.get()]
            start_dt = datetime(
                start_date.year, start_date.month, start_date.day,
                int(self.custom_start_hour_menu.get()), int(self.custom_start_min_menu.get()),
            )
            end_dt = datetime(
                end_date.year, end_date.month, end_date.day,
                int(self.custom_end_hour_menu.get()), int(self.custom_end_min_menu.get()), 59,
            )
        except (KeyError, ValueError):
            self.status_label.configure(text="Invalid custom date/time range.")
            return

        if start_dt > end_dt:
            self.status_label.configure(text="Start must be before end — swap them and try again.")
            return

        self.custom_start = start_dt
        self.custom_end = end_dt
        self.hours_window = "custom"
        save_mail_settings(self.watchlist, self.hours_window, self.custom_start, self.custom_end)
        self.refresh()

    def _notify_new_mail(self, emails):
        for e in emails:
            if not e["unread"] or e["received"] <= self.session_start or e["id"] in self.notified_ids:
                continue
            self.notified_ids.add(e["id"])
            try:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except Exception:
                pass
            self.app._show_toast(f"\U0001F4E7 {e['sender']}: {e['subject']}"[:90], ACCENT)

    def _show_error(self, message: str):
        if self.mail_tabs_container is None:
            return
        for child in self.mail_tabs_container.winfo_children():
            child.destroy()
        self.rule_tabview = None
        ctk.CTkLabel(
            self.mail_tabs_container, text=f"Couldn't reach Outlook:\n{message}",
            text_color=DANGER_TEXT, justify="left",
        ).pack(pady=40, padx=20)

    # ---- UI ---------------------------------------------------------
    def build_tab(self, root) -> None:
        header = ctk.CTkFrame(root, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(22, 8))

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(left, text="Inbox", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY).pack(anchor="w")
        self.account_label = ctk.CTkLabel(left, text="", font=("Segoe UI", 12), text_color=TEXT_SECONDARY)
        self.account_label.pack(anchor="w", pady=(2, 0))

        ctk.CTkButton(
            header, text="⟳ Refresh", fg_color=CARD, hover_color=CARD2, text_color=TEXT_PRIMARY,
            corner_radius=10, width=100, height=34, font=("Segoe UI", 12, "bold"), command=self.refresh,
        ).pack(side="right", anchor="n")

        window_row = ctk.CTkFrame(root, fg_color="transparent")
        window_row.pack(fill="x", padx=24, pady=(0, 6))
        ctk.CTkLabel(
            window_row, text="Load emails from:", font=("Segoe UI", 12), text_color=TEXT_SECONDARY,
        ).pack(side="left", padx=(0, 8))
        hours_labels = {v: k for k, v in self.HOURS_OPTIONS.items()}
        self.hours_menu = ctk.CTkOptionMenu(
            window_row, values=list(self.HOURS_OPTIONS.keys()), fg_color=ENTRY_BG,
            button_color=ACCENT, button_hover_color=ACCENT_HOVER, corner_radius=10,
            width=140, command=self.set_hours_window,
        )
        self.hours_menu.set(hours_labels.get(self.hours_window, "Last 2 hours"))
        self.hours_menu.pack(side="left")

        self._date_options = self._build_date_options()
        date_values = list(self._date_options.keys())
        hour_values = [f"{h:02d}" for h in range(24)]
        min_values = [f"{m:02d}" for m in range(60)]

        def small_menu(parent, values):
            return ctk.CTkOptionMenu(
                parent, values=values, fg_color=ENTRY_BG, button_color=ACCENT,
                button_hover_color=ACCENT_HOVER, corner_radius=8, width=88, height=30,
                font=("Segoe UI", 11),
            )

        self.custom_range_frame = ctk.CTkFrame(root, fg_color=CARD, corner_radius=10)
        # not packed here — set_hours_window shows/hides it based on the dropdown choice

        inner = ctk.CTkFrame(self.custom_range_frame, fg_color="transparent")
        inner.pack(fill="x", padx=12, pady=10)

        ctk.CTkLabel(inner, text="From", font=("Segoe UI", 11), text_color=TEXT_SECONDARY).pack(side="left", padx=(0, 6))
        self.custom_start_date_menu = small_menu(inner, date_values)
        self.custom_start_date_menu.set(date_values[0])
        self.custom_start_date_menu.pack(side="left", padx=(0, 4))
        self.custom_start_hour_menu = small_menu(inner, hour_values)
        self.custom_start_hour_menu.set("00")
        self.custom_start_hour_menu.pack(side="left", padx=(0, 2))
        ctk.CTkLabel(inner, text=":", text_color=TEXT_SECONDARY).pack(side="left")
        self.custom_start_min_menu = small_menu(inner, min_values)
        self.custom_start_min_menu.set("00")
        self.custom_start_min_menu.pack(side="left", padx=(2, 10))

        ctk.CTkLabel(inner, text="to", font=("Segoe UI", 11), text_color=TEXT_SECONDARY).pack(side="left", padx=(0, 6))
        self.custom_end_date_menu = small_menu(inner, date_values)
        self.custom_end_date_menu.set(date_values[0])
        self.custom_end_date_menu.pack(side="left", padx=(0, 4))
        self.custom_end_hour_menu = small_menu(inner, hour_values)
        self.custom_end_hour_menu.set("23")
        self.custom_end_hour_menu.pack(side="left", padx=(0, 2))
        ctk.CTkLabel(inner, text=":", text_color=TEXT_SECONDARY).pack(side="left")
        self.custom_end_min_menu = small_menu(inner, min_values)
        self.custom_end_min_menu.set("59")
        self.custom_end_min_menu.pack(side="left", padx=(2, 10))

        ctk.CTkButton(
            inner, text="Apply", fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8,
            width=70, height=30, font=("Segoe UI", 11, "bold"), command=self.apply_custom_range,
        ).pack(side="left")

        if self.hours_window == "custom":
            self.custom_range_frame.pack(fill="x", padx=24, pady=(0, 6))
            if self.custom_start:
                self.custom_start_date_menu.set(self.custom_start.strftime("%b %d, %Y"))
                self.custom_start_hour_menu.set(f"{self.custom_start.hour:02d}")
                self.custom_start_min_menu.set(f"{self.custom_start.minute:02d}")
            if self.custom_end:
                self.custom_end_date_menu.set(self.custom_end.strftime("%b %d, %Y"))
                self.custom_end_hour_menu.set(f"{self.custom_end.hour:02d}")
                self.custom_end_min_menu.set(f"{self.custom_end.minute:02d}")

        self.status_label = ctk.CTkLabel(root, text="", font=("Segoe UI", 11), text_color=TEXT_MUTED)
        self.status_label.pack(anchor="w", padx=24, pady=(0, 8))

        self.mail_tabs_container = ctk.CTkFrame(root, fg_color="transparent")
        self.mail_tabs_container.pack(fill="both", expand=True, padx=18, pady=(0, 16))

        if not OUTLOOK_AVAILABLE:
            ctk.CTkLabel(
                self.mail_tabs_container,
                text="Outlook integration isn't available (pywin32 not installed).",
                text_color=TEXT_MUTED,
            ).pack(pady=40)
            return

        self._rebuild_rule_tabs()
        self.app.root.after(200, self.refresh)

    def _rule_tab_label(self, email: str) -> str:
        return email if len(email) <= 26 else email[:23] + "…"

    def _rebuild_rule_tabs(self) -> None:
        """(Re)build one sub-tab per watched address ('rule') plus a Rules tab
        for managing the watchlist, preserving the current selection if it
        still exists after the rebuild."""
        previous_selection = None
        if self.rule_tabview is not None:
            try:
                previous_selection = self.rule_tabview.get()
            except Exception:
                previous_selection = None
            self.rule_tabview.destroy()

        self.rule_tabview = ctk.CTkTabview(
            self.mail_tabs_container,
            fg_color="transparent",
            segmented_button_fg_color=CARD,
            segmented_button_selected_color=ACCENT,
            segmented_button_selected_hover_color=ACCENT_HOVER,
            segmented_button_unselected_color=CARD,
            segmented_button_unselected_hover_color=CARD2,
            text_color=TEXT_PRIMARY,
        )
        self.rule_tabview.pack(fill="both", expand=True)

        rules_tab = self.rule_tabview.add(self.RULES_TAB_NAME)
        self._build_rules_tab(rules_tab)

        self.rule_frames = {}
        self.rule_tab_labels = {}
        used_labels = set()
        for email in self.watchlist:
            label = self._rule_tab_label(email)
            base_label = label
            n = 2
            while label in used_labels:
                label = f"{base_label} ({n})"
                n += 1
            used_labels.add(label)
            self.rule_tab_labels[email] = label

            tab = self.rule_tabview.add(label)
            header = ctk.CTkLabel(
                tab, text=email, font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY,
            )
            header.pack(anchor="w", padx=8, pady=(8, 4))
            frame = ctk.CTkScrollableFrame(tab, fg_color="transparent")
            frame.pack(fill="both", expand=True, padx=2)
            self.rule_frames[email] = frame

        self._built_rule_keys = list(self.watchlist)

        self.folder_frames = {}
        folder_labels = set()
        for folder_name in self.rule_folders:
            label = f"📁 {folder_name}"
            folder_labels.add(label)
            tab = self.rule_tabview.add(label)
            ctk.CTkLabel(
                tab, text=f"Filed here by your Outlook rule(s) — folder \"{folder_name}\"",
                font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY,
            ).pack(anchor="w", padx=8, pady=(8, 4))
            frame = ctk.CTkScrollableFrame(tab, fg_color="transparent")
            frame.pack(fill="both", expand=True, padx=2)
            self.folder_frames[folder_name] = frame
        self._built_folder_keys = list(self.rule_folders.keys())

        target = (
            previous_selection
            if previous_selection in ([self.RULES_TAB_NAME] + list(used_labels) + list(folder_labels))
            else self.RULES_TAB_NAME
        )
        try:
            self.rule_tabview.set(target)
        except Exception:
            pass

    def _build_rules_tab(self, tab) -> None:
        watch_row = ctk.CTkFrame(tab, fg_color="transparent")
        watch_row.pack(fill="x", padx=8, pady=(12, 6))
        self.watch_entry = ctk.CTkEntry(
            watch_row, placeholder_text="Add an email or group name to monitor...", fg_color=ENTRY_BG,
            border_color=BORDER, border_width=1, corner_radius=10, height=36, font=("Segoe UI", 12),
        )
        self.watch_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.watch_entry.bind("<Return>", lambda _e: self.add_watch_email())
        self.app._floating_circle_button(
            watch_row, "+", 36, ACCENT, ACCENT_HOVER, "white", ("Segoe UI", 16, "bold"), self.add_watch_email
        ).pack(side="left")

        ctk.CTkLabel(
            tab, text="Each address below gets its own tab showing only the mail matching it.",
            font=("Segoe UI", 11), text_color=TEXT_MUTED,
        ).pack(anchor="w", padx=8, pady=(0, 8))

        self.chips_frame = ctk.CTkFrame(tab, fg_color="transparent")
        self.chips_frame.pack(fill="x", padx=8)
        self._render_chips()

    def _render_chips(self) -> None:
        if self.chips_frame is None:
            return
        for child in self.chips_frame.winfo_children():
            child.destroy()
        if not self.watchlist:
            ctk.CTkLabel(
                self.chips_frame, text="Not monitoring any addresses yet — add one above.",
                font=("Segoe UI", 11), text_color=TEXT_MUTED,
            ).pack(anchor="w")
            return
        for email in self.watchlist:
            chip = ctk.CTkFrame(self.chips_frame, fg_color=CARD, corner_radius=8)
            chip.pack(anchor="w", pady=3)
            ctk.CTkLabel(
                chip, text=email, font=("Segoe UI", 11), text_color=TEXT_PRIMARY,
            ).pack(side="left", padx=(10, 6), pady=4)
            ctk.CTkButton(
                chip, text="✕", width=18, height=18, corner_radius=6, fg_color="transparent",
                hover_color=DANGER, text_color=TEXT_SECONDARY, font=("Segoe UI", 10),
                command=lambda em=email: self.remove_watch_email(em),
            ).pack(side="left", padx=(0, 8))

    def _render(self, current_user) -> None:
        self.account_label.configure(text=f"Signed in as {current_user}" if current_user else "")
        unread = sum(1 for e in self.emails if e["unread"])
        folder_note = f" · {len(self.rule_folders)} rule-folder tab(s)" if self.rule_folders else ""
        if self.watchlist or self.rule_folders:
            if self.hours_window == "custom" and self.custom_start and self.custom_end:
                window_label = (
                    f"{self.custom_start.strftime('%b %d %I:%M %p')} → "
                    f"{self.custom_end.strftime('%b %d %I:%M %p')}"
                )
            else:
                hours_labels = {v: k for k, v in self.HOURS_OPTIONS.items()}
                window_label = hours_labels.get(self.hours_window, "Last 2 hours")
            self.status_label.configure(
                text=(
                    f"Monitoring {len(self.watchlist)} address(es) · {window_label} · {unread} unread · "
                    f"showing latest {len(self.emails)} · last refreshed {fmt_12h(datetime.now().isoformat())}"
                    f"{folder_note}"
                )
            )
        else:
            self.status_label.configure(text="")

        if (
            self.rule_tabview is None
            or sorted(self.watchlist) != sorted(self._built_rule_keys)
            or sorted(self.rule_folders.keys()) != sorted(self._built_folder_keys)
        ):
            self._rebuild_rule_tabs()
        else:
            self._render_chips()

        for email, frame in self.rule_frames.items():
            for child in frame.winfo_children():
                child.destroy()
            matches = [e for e in self.emails if email in e.get("matched_rules", [])]
            if not matches:
                ctk.CTkLabel(
                    frame, text="No recent messages matching this address.", text_color=TEXT_MUTED,
                ).pack(pady=30)
                continue
            unread_here = sum(1 for e in matches if e["unread"])
            ctk.CTkLabel(
                frame, text=f"{len(matches)} message(s) · {unread_here} unread",
                font=("Segoe UI", 10), text_color=TEXT_MUTED,
            ).pack(anchor="w", padx=6, pady=(0, 4))
            for e in matches:
                self._build_row(frame, e)

        for folder_name, frame in self.folder_frames.items():
            for child in frame.winfo_children():
                child.destroy()
            matches = self.folder_mail.get(folder_name, [])
            if not matches:
                ctk.CTkLabel(
                    frame, text="No recent messages in this folder.", text_color=TEXT_MUTED,
                ).pack(pady=30)
                continue
            unread_here = sum(1 for e in matches if e["unread"])
            ctk.CTkLabel(
                frame, text=f"{len(matches)} message(s) · {unread_here} unread",
                font=("Segoe UI", 10), text_color=TEXT_MUTED,
            ).pack(anchor="w", padx=6, pady=(0, 4))
            for e in matches:
                self._build_row(frame, e)

    def _format_received(self, dt) -> str:
        today = datetime.now().date()
        if dt.date() == today:
            return fmt_12h(dt.isoformat())
        if dt.date() == today - timedelta(days=1):
            return f"Yesterday, {fmt_12h(dt.isoformat())}"
        return f"{dt.strftime('%b %d')}, {fmt_12h(dt.isoformat())}"

    def _build_row(self, parent, e: dict) -> None:
        row = ctk.CTkFrame(
            parent, fg_color=CARD2 if e["unread"] else "transparent", corner_radius=10
        )
        row.pack(fill="x", pady=4, padx=2)
        click_targets = [row]

        stripe_color = PRIORITY_COLORS["high"] if e["importance_high"] else (ACCENT if e["unread"] else BORDER)
        stripe = ctk.CTkFrame(row, width=4, corner_radius=0, fg_color=stripe_color)
        stripe.pack(side="left", fill="y")
        stripe.pack_propagate(False)

        content = ctk.CTkFrame(row, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True, padx=(10, 10), pady=8)
        click_targets.append(content)

        top_line = ctk.CTkFrame(content, fg_color="transparent")
        top_line.pack(fill="x")
        click_targets.append(top_line)

        weight = "bold" if e["unread"] else "normal"

        sender_side = ctk.CTkFrame(top_line, fg_color="transparent")
        sender_side.pack(side="left")
        click_targets.append(sender_side)

        if e.get("folder") == "sent":
            ctk.CTkLabel(
                sender_side, text="📤", font=("Segoe UI", 10),
            ).pack(side="left", padx=(0, 4))

        sender_lbl = ctk.CTkLabel(
            sender_side, text=e["sender"], font=("Segoe UI", 12, weight), text_color=TEXT_PRIMARY, anchor="w"
        )
        sender_lbl.pack(side="left")
        click_targets.append(sender_lbl)

        if e["unread"]:
            ctk.CTkLabel(
                sender_side, text="●", font=("Segoe UI", 8), text_color=ACCENT,
            ).pack(side="left", padx=(6, 0))

        time_lbl = ctk.CTkLabel(
            top_line, text=self._format_received(e["received"]), font=("Segoe UI", 10), text_color=TEXT_MUTED
        )
        time_lbl.pack(side="right")
        click_targets.append(time_lbl)

        subject_line = ctk.CTkFrame(content, fg_color="transparent")
        subject_line.pack(fill="x", pady=(3, 0))
        click_targets.append(subject_line)

        subject_lbl = ctk.CTkLabel(
            subject_line, text=e["subject"], font=("Segoe UI", 12, weight), text_color=TEXT_PRIMARY,
            anchor="w", justify="left", wraplength=680,
        )
        subject_lbl.pack(side="left", fill="x")
        click_targets.append(subject_lbl)

        flags = []
        if e["importance_high"]:
            flags.append("❗ High")
        if e["attachment_count"]:
            flags.append(f"📎 {e['attachment_count']}")
        if flags:
            flags_lbl = ctk.CTkLabel(
                subject_line, text="   " + "   ".join(flags), font=("Segoe UI", 10, "bold"),
                text_color=PRIORITY_COLORS["high"] if e["importance_high"] else TEXT_SECONDARY,
            )
            flags_lbl.pack(side="left")
            click_targets.append(flags_lbl)

        if e.get("to_names") and (e.get("folder") == "sent" or len(e["to_names"]) > 1):
            to_lbl = ctk.CTkLabel(
                content, text="To: " + ", ".join(e["to_names"][:4]) + (" +more" if len(e["to_names"]) > 4 else ""),
                font=("Segoe UI", 10), text_color=TEXT_MUTED, anchor="w", justify="left", wraplength=760,
            )
            to_lbl.pack(fill="x", anchor="w", pady=(2, 0))
            click_targets.append(to_lbl)

        if e["preview"]:
            preview_lbl = ctk.CTkLabel(
                content, text=e["preview"], font=("Segoe UI", 11), text_color=TEXT_SECONDARY,
                anchor="w", justify="left", wraplength=760,
            )
            preview_lbl.pack(fill="x", anchor="w", pady=(3, 0))
            click_targets.append(preview_lbl)

        for w in click_targets:
            w.bind("<Button-1>", lambda _e, mid=e["id"]: self.open_in_outlook(mid))
            w.configure(cursor="hand2")

    def open_in_outlook(self, entry_id: str) -> None:
        try:
            ns = self._get_namespace()
            item = ns.GetItemFromID(entry_id)
            item.Display()
        except Exception as e:
            messagebox.showerror("Couldn't open message", str(e), parent=self.app.root)


class PortalBrowserManager:
    """Manages a list of saved portal shortcuts (name + URL) and launches a
    real, persistent, navigable embedded browser window for any of them —
    or any one-off URL — via portal_browser_helper.py.

    Runs as a subprocess: the browser engine needs the main thread, which
    this app's Tkinter loop already owns, so this can't literally be a tab's
    content the way the board/mail tabs are — it's a separate window the app
    launches and leaves running independently.

    Login sessions with an explicit-expiry cookie persist across launches
    (verified: set with a real expiry in one process, read back correctly in
    a separate one) via a shared profile directory. Session-only cookies —
    common for SSO — won't survive a restart; that's normal browser
    behavior, not a limitation here.
    """

    def __init__(self, app):
        self.app = app
        self.portals: list = load_portals()
        self.list_frame = None
        self.name_entry = None
        self.url_entry = None

    def _save(self) -> None:
        save_portals(self.portals)

    def add_portal(self) -> None:
        name = self.name_entry.get().strip()
        url = self.url_entry.get().strip()
        if not name or not url:
            return
        if "://" not in url:
            url = "https://" + url
        self.portals.append({"name": name, "url": url})
        self._save()
        self.name_entry.delete(0, "end")
        self.url_entry.delete(0, "end")
        self._render_list()

    def remove_portal(self, index: int) -> None:
        if 0 <= index < len(self.portals):
            del self.portals[index]
            self._save()
            self._render_list()

    def open_url(self, url: str) -> None:
        url = url.strip()
        if not url:
            return
        if "://" not in url:
            url = "https://" + url
        if not PORTAL_BROWSER_HELPER.exists():
            messagebox.showerror(
                "Missing file", "portal_browser_helper.py must be alongside todo_app.py.",
                parent=self.app.root,
            )
            return
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        exe = str(pythonw) if pythonw.exists() else sys.executable
        PORTAL_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.Popen(
                [exe, str(PORTAL_BROWSER_HELPER), url, str(PORTAL_PROFILE_DIR)],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as e:
            messagebox.showerror("Couldn't open browser window", str(e), parent=self.app.root)

    def open_in_native_browser(self, url: str) -> None:
        """Opens the URL in the user's actual default system browser (their
        real Edge/Chrome, with their normal profile, extensions, saved
        passwords) rather than the app's embedded browser window."""
        url = url.strip()
        if not url:
            return
        if "://" not in url:
            url = "https://" + url
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Couldn't open browser", str(e), parent=self.app.root)

    # ---- UI ---------------------------------------------------------
    def build_tab(self, root) -> None:
        header = ctk.CTkFrame(root, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(22, 8))
        ctk.CTkLabel(
            header, text="Portals", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Opens a real, persistent browser window — log in and use any portal normally.",
            font=("Segoe UI", 12), text_color=TEXT_SECONDARY,
        ).pack(anchor="w", pady=(2, 0))

        quick_row = ctk.CTkFrame(root, fg_color="transparent")
        quick_row.pack(fill="x", padx=24, pady=(14, 6))
        quick_entry = ctk.CTkEntry(
            quick_row, placeholder_text="Enter a URL to open once...", fg_color=ENTRY_BG,
            border_color=BORDER, border_width=1, corner_radius=10, height=38, font=("Segoe UI", 13),
        )
        quick_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        quick_entry.bind("<Return>", lambda _e: self.open_url(quick_entry.get()))
        ctk.CTkButton(
            quick_row, text="Open", fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=10,
            width=90, height=38, command=lambda: self.open_url(quick_entry.get()),
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            quick_row, text="🌐 Native Browser", fg_color=CARD, hover_color=CARD2, text_color=TEXT_PRIMARY,
            corner_radius=10, width=140, height=38, command=lambda: self.open_in_native_browser(quick_entry.get()),
        ).pack(side="left")

        ctk.CTkLabel(
            root, text="Saved portals", font=("Segoe UI", 13, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=24, pady=(16, 4))

        add_row = ctk.CTkFrame(root, fg_color="transparent")
        add_row.pack(fill="x", padx=24, pady=(0, 10))
        self.name_entry = ctk.CTkEntry(
            add_row, placeholder_text="Name", fg_color=ENTRY_BG, border_color=BORDER,
            border_width=1, corner_radius=10, height=36, width=160, font=("Segoe UI", 12),
        )
        self.name_entry.pack(side="left", padx=(0, 8))
        self.url_entry = ctk.CTkEntry(
            add_row, placeholder_text="URL", fg_color=ENTRY_BG, border_color=BORDER,
            border_width=1, corner_radius=10, height=36, font=("Segoe UI", 12),
        )
        self.url_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.url_entry.bind("<Return>", lambda _e: self.add_portal())
        self.app._floating_circle_button(
            add_row, "+", 36, ACCENT, ACCENT_HOVER, "white", ("Segoe UI", 16, "bold"), self.add_portal
        ).pack(side="left")

        self.list_frame = ctk.CTkScrollableFrame(root, fg_color="transparent")
        self.list_frame.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        self._render_list()

    def _render_list(self) -> None:
        if self.list_frame is None:
            return
        for child in self.list_frame.winfo_children():
            child.destroy()
        if not self.portals:
            ctk.CTkLabel(
                self.list_frame, text="No saved portals yet — add one above.", text_color=TEXT_MUTED,
            ).pack(pady=30)
            return
        for i, portal in enumerate(self.portals):
            row = ctk.CTkFrame(self.list_frame, fg_color=CARD2, corner_radius=10)
            row.pack(fill="x", pady=4, padx=2)

            content = ctk.CTkFrame(row, fg_color="transparent")
            content.pack(side="left", fill="both", expand=True, padx=(12, 8), pady=10)
            ctk.CTkLabel(
                content, text=portal["name"], font=("Segoe UI", 13, "bold"), text_color=TEXT_PRIMARY,
                anchor="w",
            ).pack(fill="x", anchor="w")
            ctk.CTkLabel(
                content, text=portal["url"], font=("Segoe UI", 11), text_color=TEXT_SECONDARY,
                anchor="w",
            ).pack(fill="x", anchor="w")

            ctk.CTkButton(
                row, text="✕", fg_color="transparent", hover_color=DANGER, text_color=TEXT_SECONDARY,
                corner_radius=8, width=32, height=32, command=lambda idx=i: self.remove_portal(idx),
            ).pack(side="right", padx=(8, 4))
            ctk.CTkButton(
                row, text="🌐 Native", fg_color=CARD, hover_color=CARD2, text_color=TEXT_PRIMARY,
                corner_radius=8, width=90, height=32,
                command=lambda u=portal["url"]: self.open_in_native_browser(u),
            ).pack(side="right", padx=(0, 6))
            ctk.CTkButton(
                row, text="Open", fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8,
                width=70, height=32, command=lambda u=portal["url"]: self.open_url(u),
            ).pack(side="right", padx=(0, 6))


class TodoApp:
    def __init__(self):
        self.data = load_data()
        self.action_queue: "queue.Queue[str]" = queue.Queue()
        self.card_widgets: dict = {}
        self.column_frames: dict = {}
        self.column_containers: dict = {}
        self.column_headers: dict = {}
        self.drag_data = {"task_id": None, "start_x": 0, "start_y": 0, "ghost": None, "dragging": False}

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.root = ctk.CTk()
        self.root.title("Daily Todo")
        self.root.geometry("980x640")
        self.root.minsize(760, 480)
        self.root.configure(fg_color=BG)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        self.strike_font = ctk.CTkFont(family="Segoe UI", size=13, overstrike=True)
        self.normal_font = ctk.CTkFont(family="Segoe UI", size=13, overstrike=False)

        self.notes_data = load_notes()
        self.notes_list_frame = None
        self.active_notifications: list = []

        self.weekly_board = PeriodTaskBoard(
            self, "Add a task for this week...", WEEKLY_FILE, WEEKLY_HISTORY_FILE,
            week_key, week_label, self.strike_font, self.normal_font,
        )
        self.half_year_board = PeriodTaskBoard(
            self, "Add a task for this half-year...", HALF_YEAR_FILE, HALF_YEAR_HISTORY_FILE,
            half_year_key, half_year_label, self.strike_font, self.normal_font,
        )
        self.yearly_board = PeriodTaskBoard(
            self, "Add a task for this year...", YEARLY_FILE, YEARLY_HISTORY_FILE,
            year_key, year_label, self.strike_font, self.normal_font,
        )
        self.mail_monitor = MailMonitor(self)
        self.portal_browser = PortalBrowserManager(self)

        self._build_ui()
        self.render()
        self.render_notes()

        self.tray_icon = None
        self._start_tray()
        self.root.after(150, self._poll_queue)
        self.root.after(5000, self._check_schedules)

    # ---- UI construction ---------------------------------------------
    def _build_ui(self):
        self.tabview = ctk.CTkTabview(
            self.root,
            fg_color=BG,
            segmented_button_fg_color=CARD,
            segmented_button_selected_color=ACCENT,
            segmented_button_selected_hover_color=ACCENT_HOVER,
            segmented_button_unselected_color=CARD,
            segmented_button_unselected_hover_color=CARD2,
            text_color=TEXT_PRIMARY,
        )
        self.tabview.pack(fill="both", expand=True, padx=0, pady=0)
        board_tab = self.tabview.add("📋 Board")
        weekly_tab = self.tabview.add("📅 Weekly")
        half_year_tab = self.tabview.add("🗓️ 6-Month")
        yearly_tab = self.tabview.add("📆 Yearly")
        notes_tab = self.tabview.add("📝 Notes")
        mail_tab = self.tabview.add("📧 Mail")
        portals_tab = self.tabview.add("🌐 Portals")

        self._build_board_tab(board_tab)
        self.weekly_board.build_tab(weekly_tab)
        self.half_year_board.build_tab(half_year_tab)
        self.yearly_board.build_tab(yearly_tab)
        self._build_notes_tab(notes_tab)
        self.mail_monitor.build_tab(mail_tab)
        self.portal_browser.build_tab(portals_tab)

    def _floating_circle_button(self, parent, text, diameter, bg, hover, text_color, font, command):
        """A round, floating-looking button.

        NOTE: CustomTkinter has a real rendering bug where corner_radius == diameter/2
        (a true circle) makes the button silently render far wider than requested when
        Windows display scaling is active (reproduced at 150% scaling: a 44px button
        balloons to ~87px wide). That previously made this button un-clickable. Capping
        the radius at ~30% of the diameter stays well clear of the bug while still
        reading as a soft, pill-shaped "floating" button.
        """
        radius = max(4, int(diameter * 0.3))
        btn = ctk.CTkButton(
            parent,
            width=diameter,
            height=diameter,
            corner_radius=radius,
            fg_color=bg,
            hover_color=hover,
            text=text,
            text_color=text_color,
            font=font,
            border_width=1,
            border_color=SHADOW_COLOR,
            command=command,
        )
        return btn

    def _build_board_tab(self, root):
        header = ctk.CTkFrame(root, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(22, 8))

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)

        self.date_label = ctk.CTkLabel(
            left, text="", font=("Segoe UI", 22, "bold"), text_color=TEXT_PRIMARY
        )
        self.date_label.pack(anchor="w")

        self.subtitle_label = ctk.CTkLabel(
            left, text="", font=("Segoe UI", 13), text_color=TEXT_SECONDARY
        )
        self.subtitle_label.pack(anchor="w", pady=(2, 0))

        history_btn = ctk.CTkButton(
            header,
            text="🕘 History",
            fg_color=CARD,
            hover_color=CARD2,
            text_color=TEXT_PRIMARY,
            corner_radius=10,
            width=110,
            height=34,
            font=("Segoe UI", 12, "bold"),
            command=self.open_history,
        )
        history_btn.pack(side="right", anchor="n")

        progress_frame = ctk.CTkFrame(root, fg_color="transparent")
        progress_frame.pack(fill="x", padx=24, pady=(2, 4))

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, height=10, corner_radius=6, fg_color=CARD, progress_color=ACCENT
        )
        self.progress_bar.pack(fill="x", side="left", expand=True)
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_frame, text="0%", font=("Segoe UI", 12, "bold"), text_color=ACCENT, width=40
        )
        self.progress_label.pack(side="left", padx=(10, 0))

        entry_frame = ctk.CTkFrame(root, fg_color="transparent")
        entry_frame.pack(fill="x", padx=24, pady=(16, 10))

        self.entry = ctk.CTkEntry(
            entry_frame,
            placeholder_text="Add a task for today...",
            fg_color=ENTRY_BG,
            border_color=BORDER,
            border_width=1,
            corner_radius=12,
            height=42,
            font=("Segoe UI", 13),
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry.bind("<Return>", lambda _e: self.add_task())

        add_btn = self._floating_circle_button(
            entry_frame, "+", 44, ACCENT, ACCENT_HOVER, "white", ("Segoe UI", 20, "bold"), self.add_task
        )
        add_btn.pack(side="left")

        board = ctk.CTkFrame(root, fg_color="transparent")
        board.pack(fill="both", expand=True, padx=18, pady=(4, 4))
        board.grid_columnconfigure(0, weight=1, uniform="col")
        board.grid_columnconfigure(1, weight=1, uniform="col")
        board.grid_columnconfigure(2, weight=1, uniform="col")
        board.grid_rowconfigure(0, weight=1)

        for i, status in enumerate(STATUS_ORDER):
            container = ctk.CTkFrame(board, fg_color=CARD, corner_radius=14, border_width=0, border_color=ACCENT)
            container.grid(row=0, column=i, sticky="nsew", padx=6)
            container.grid_rowconfigure(1, weight=1)
            container.grid_columnconfigure(0, weight=1)
            self.column_containers[status] = container

            col_header = ctk.CTkFrame(container, fg_color="transparent")
            col_header.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))

            dot = ctk.CTkFrame(
                col_header, width=9, height=9, corner_radius=5, fg_color=COLUMN_ACCENTS[status]
            )
            dot.pack(side="left", padx=(0, 8))
            dot.pack_propagate(False)

            ctk.CTkLabel(
                col_header,
                text=STATUS_LABELS[status],
                font=("Segoe UI", 13, "bold"),
                text_color=TEXT_PRIMARY,
            ).pack(side="left")

            count_lbl = ctk.CTkLabel(
                col_header, text="0", font=("Segoe UI", 12), text_color=TEXT_SECONDARY
            )
            count_lbl.pack(side="right")
            self.column_headers[status] = count_lbl

            scroll = ctk.CTkScrollableFrame(
                container, fg_color="transparent", scrollbar_button_color=BORDER
            )
            scroll.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 12))
            self.column_frames[status] = scroll

        footer = ctk.CTkFrame(root, fg_color="transparent")
        footer.pack(fill="x", padx=24, pady=(4, 18))

        self.count_label = ctk.CTkLabel(
            footer, text="", font=("Segoe UI", 12), text_color=TEXT_SECONDARY
        )
        self.count_label.pack(side="left")

        clear_btn = ctk.CTkButton(
            footer,
            text="Clear completed",
            fg_color="transparent",
            hover_color=CARD,
            text_color=TEXT_SECONDARY,
            font=("Segoe UI", 12),
            width=120,
            height=26,
            corner_radius=8,
            command=self.clear_completed,
        )
        clear_btn.pack(side="right")

        hint = ctk.CTkLabel(
            footer,
            text="Drag a card between columns · click a card for details & notes",
            font=("Segoe UI", 11),
            text_color=TEXT_MUTED,
        )
        hint.pack(side="right", padx=(0, 16))

    def _build_notes_tab(self, root):
        top_bar = ctk.CTkFrame(root, fg_color="transparent")
        top_bar.pack(fill="x", padx=24, pady=(20, 10))

        ctk.CTkLabel(
            top_bar, text="Notes", font=("Segoe UI", 20, "bold"), text_color=TEXT_PRIMARY
        ).pack(side="left")

        ctk.CTkLabel(
            top_bar,
            text="Logins, company info, anything worth keeping around",
            font=("Segoe UI", 12),
            text_color=TEXT_SECONDARY,
        ).pack(side="left", padx=(10, 0))

        ctk.CTkButton(
            top_bar,
            text="+ New Note",
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            corner_radius=10,
            width=120,
            height=32,
            font=("Segoe UI", 12, "bold"),
            command=lambda: self._open_note_dialog(None),
        ).pack(side="right")

        self.notes_list_frame = ctk.CTkScrollableFrame(root, fg_color="transparent")
        self.notes_list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 16))

    # ---- Data operations ------------------------------------------------
    def _find_task(self, task_id):
        for t in self.data["tasks"]:
            if t["id"] == task_id:
                return t
        return None

    def add_task(self, text: str = None, parent_id: int = None):
        """Adds a task. With no args, reads from the main quick-add entry (the
        normal "+"-button flow). Passing text/parent_id directly is how a
        subtask gets created from the task dialog — it's a fully independent
        task (its own card, its own status/priority), just tagged with which
        task it's linked to."""
        from_entry = text is None
        if from_entry:
            text = self.entry.get().strip()
        else:
            text = text.strip()
        if not text:
            return None
        todo_orders = [t.get("order", 0) for t in self.data["tasks"] if t["status"] == "todo"]
        task = {
            "id": self.data["next_id"],
            "text": text,
            "notes": "",
            "status": "todo",
            "priority": "medium",
            "order": (max(todo_orders) + 1) if todo_orders else 0,
            "created_at": datetime.now().isoformat(),
            "started_at": None,
            "completed_at": None,
            "start_time": None,
            "end_time": None,
            "reminded_for_end_time": None,
            "parent_id": parent_id,
        }
        self.data["next_id"] += 1
        self.data["tasks"].append(task)
        if from_entry:
            self.entry.delete(0, "end")
        save_data(self.data)
        self.render()
        return task["id"]

    def delete_task(self, task_id: int):
        task = self._find_task(task_id)
        if task is not None and task["status"] == "done":
            record_completed_task(task)
        for t in self.data["tasks"]:
            if t.get("parent_id") == task_id:
                t["parent_id"] = None  # subtasks are independent tasks — unlink, don't delete
        self.data["tasks"] = [t for t in self.data["tasks"] if t["id"] != task_id]
        save_data(self.data)
        self.render()

    def clear_completed(self):
        for t in self.data["tasks"]:
            if t["status"] == "done":
                record_completed_task(t)
        self.data["tasks"] = [t for t in self.data["tasks"] if t["status"] != "done"]
        save_data(self.data)
        self.render()

    # ---- Rendering --------------------------------------------------------
    def render(self):
        self.date_label.configure(text=datetime.now().strftime("%A, %B %d"))

        tasks = self.data["tasks"]
        total = len(tasks)
        done = sum(1 for t in tasks if t["status"] == "done")

        self.subtitle_label.configure(
            text="All clear for today ✨" if total == 0 else f"{total - done} task(s) remaining"
        )
        self.progress_bar.set(done / total if total else 0)
        self.progress_label.configure(text=f"{round(done / total * 100) if total else 0}%")
        self.count_label.configure(text=f"{done} of {total} completed")

        self.card_widgets = {}
        for status in STATUS_ORDER:
            frame = self.column_frames[status]
            for child in frame.winfo_children():
                child.destroy()

            subset = [t for t in tasks if t["status"] == status]
            self.column_headers[status].configure(text=str(len(subset)))

            if not subset:
                ctk.CTkLabel(
                    frame, text="Drop tasks here", font=("Segoe UI", 11), text_color=TEXT_MUTED
                ).pack(pady=18)
                continue

            for priority in PRIORITY_ORDER:
                group = sorted(
                    [t for t in subset if t.get("priority", "medium") == priority],
                    key=lambda t: t.get("order", 0),
                )
                if not group:
                    continue

                group_header = ctk.CTkFrame(frame, fg_color="transparent")
                group_header.pack(fill="x", pady=(10, 2))
                ctk.CTkLabel(
                    group_header, text="", width=8, height=8, corner_radius=4,
                    fg_color=PRIORITY_COLORS[priority],
                ).pack(side="left", padx=(2, 6))
                ctk.CTkLabel(
                    group_header, text=f"{PRIORITY_LABELS[priority]} priority",
                    font=("Segoe UI", 10, "bold"), text_color=TEXT_MUTED,
                ).pack(side="left")

                for t in group:
                    self.card_widgets[t["id"]] = self._build_card(frame, t)

    def _build_card(self, parent, task: dict):
        status = task["status"]
        card = ctk.CTkFrame(
            parent, fg_color=CARD_DONE if status == "done" else CARD2, corner_radius=10
        )
        card.pack(fill="x", pady=5, padx=2)

        stripe = ctk.CTkFrame(
            card, width=4, corner_radius=0, fg_color=COLUMN_ACCENTS[status]
        )
        stripe.pack(side="left", fill="y")
        stripe.pack_propagate(False)

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True, padx=(10, 4), pady=8)

        priority = task.get("priority", "medium")
        badge = ctk.CTkLabel(
            content,
            text=PRIORITY_LABELS[priority],
            font=("Segoe UI", 9, "bold"),
            text_color=PRIORITY_TEXT_COLORS[priority],
            fg_color=PRIORITY_COLORS[priority],
            corner_radius=6,
        )
        badge.pack(anchor="w", pady=(0, 5), ipadx=6, ipady=1)

        title_lbl = ctk.CTkLabel(
            content,
            text=task["text"],
            font=self.strike_font if status == "done" else self.normal_font,
            text_color=TEXT_MUTED if status == "done" else TEXT_PRIMARY,
            anchor="w",
            justify="left",
            wraplength=185,
        )
        title_lbl.pack(fill="x", anchor="w")

        drag_targets = [card, content, badge, title_lbl]

        if task.get("notes"):
            preview = task["notes"].splitlines()[0][:70]
            note_lbl = ctk.CTkLabel(
                content,
                text=f"\U0001F4DD {preview}",
                font=("Segoe UI", 10),
                text_color=TEXT_SECONDARY,
                anchor="w",
                justify="left",
                wraplength=185,
            )
            note_lbl.pack(fill="x", anchor="w", pady=(4, 0))
            drag_targets.append(note_lbl)

        if task.get("start_time") or task.get("end_time"):
            schedule_text = "🕐 " + fmt_12h(task.get("start_time"))
            if task.get("end_time"):
                schedule_text += f" – {fmt_12h(task.get('end_time'))}"
            schedule_lbl = ctk.CTkLabel(
                content, text=schedule_text, font=("Segoe UI", 10), text_color=TEXT_SECONDARY,
                anchor="w", justify="left", wraplength=185,
            )
            schedule_lbl.pack(fill="x", anchor="w", pady=(4, 0))
            drag_targets.append(schedule_lbl)

        parent_id = task.get("parent_id")
        if parent_id is not None:
            parent_task = self._find_task(parent_id)
            if parent_task is not None:
                link_lbl = ctk.CTkLabel(
                    content, text=f"↳ Subtask of: {parent_task['text'][:40]}",
                    font=("Segoe UI", 10), text_color=ACCENT, anchor="w", justify="left", wraplength=185,
                )
                link_lbl.pack(fill="x", anchor="w", pady=(4, 0))
                drag_targets.append(link_lbl)

        child_count = sum(1 for t in self.data["tasks"] if t.get("parent_id") == task["id"])
        if child_count:
            children = [t for t in self.data["tasks"] if t.get("parent_id") == task["id"]]
            done_children = sum(1 for t in children if t["status"] == "done")
            subtasks_lbl = ctk.CTkLabel(
                content, text=f"\U0001F517 {done_children}/{child_count} subtask(s) done",
                font=("Segoe UI", 10), text_color=TEXT_SECONDARY, anchor="w", justify="left", wraplength=185,
            )
            subtasks_lbl.pack(fill="x", anchor="w", pady=(4, 0))
            drag_targets.append(subtasks_lbl)

        del_btn = self._floating_circle_button(
            card, "✕", 22, CARD2, DANGER, TEXT_SECONDARY, ("Segoe UI", 11),
            lambda tid=task["id"]: self.delete_task(tid),
        )
        del_btn.pack(side="right", padx=(0, 8), pady=8, anchor="n")

        for widget in drag_targets:
            self._bind_drag(widget, task["id"])
            widget.configure(cursor="hand2")

        return card

    # ---- Drag and drop -----------------------------------------------
    def _bind_drag(self, widget, task_id: int):
        widget.bind("<ButtonPress-1>", lambda e, tid=task_id: self._on_drag_start(e, tid))
        widget.bind("<B1-Motion>", self._on_drag_motion)
        widget.bind("<ButtonRelease-1>", self._on_drag_release)

    def _on_drag_start(self, event, task_id: int):
        self.drag_data = {
            "task_id": task_id,
            "start_x": event.x_root,
            "start_y": event.y_root,
            "ghost": None,
            "dragging": False,
        }

    def _on_drag_motion(self, event):
        if self.drag_data["task_id"] is None:
            return
        dx = abs(event.x_root - self.drag_data["start_x"])
        dy = abs(event.y_root - self.drag_data["start_y"])
        if not self.drag_data["dragging"] and (dx > 6 or dy > 6):
            self.drag_data["dragging"] = True
            self._create_ghost()
        if self.drag_data["dragging"] and self.drag_data["ghost"]:
            self.drag_data["ghost"].geometry(f"+{event.x_root + 14}+{event.y_root + 12}")
            self._highlight_drop_target(event)

    def _create_ghost(self):
        task = self._find_task(self.drag_data["task_id"])
        if task is None:
            return
        ghost = ctk.CTkToplevel(self.root)
        ghost.overrideredirect(True)
        ghost.attributes("-topmost", True)
        try:
            ghost.attributes("-alpha", 0.88)
        except Exception:
            pass
        ctk.CTkLabel(
            ghost,
            text=task["text"],
            fg_color=ACCENT,
            text_color="white",
            corner_radius=10,
            padx=14,
            pady=10,
            font=("Segoe UI", 12, "bold"),
        ).pack()
        self.drag_data["ghost"] = ghost

    def _column_at(self, x_root: int, y_root: int):
        for status, container in self.column_containers.items():
            fx, fy = container.winfo_rootx(), container.winfo_rooty()
            fw, fh = container.winfo_width(), container.winfo_height()
            if fx <= x_root <= fx + fw and fy <= y_root <= fy + fh:
                return status
        return None

    def _highlight_drop_target(self, event):
        target = self._column_at(event.x_root, event.y_root)
        for status, container in self.column_containers.items():
            container.configure(border_width=2 if status == target else 0)

    def _on_drag_release(self, event):
        task_id = self.drag_data.get("task_id")
        dragging = self.drag_data.get("dragging")
        ghost = self.drag_data.get("ghost")

        if ghost:
            ghost.destroy()
        for container in self.column_containers.values():
            container.configure(border_width=0)

        if task_id is None:
            self.drag_data = {"task_id": None, "start_x": 0, "start_y": 0, "ghost": None, "dragging": False}
            return

        if dragging:
            target_status = self._column_at(event.x_root, event.y_root)
            if target_status:
                self._move_task(task_id, target_status, event.y_root)
        else:
            self._open_task_dialog(task_id)

        self.drag_data = {"task_id": None, "start_x": 0, "start_y": 0, "ghost": None, "dragging": False}

    def _move_task(self, task_id: int, target_status: str, drop_y_root: int = 10 ** 9):
        task = self._find_task(task_id)
        if task is None:
            return
        old_status = task["status"]

        siblings = [t for t in self.data["tasks"] if t["status"] == target_status and t["id"] != task_id]
        positions = []
        for s in siblings:
            w = self.card_widgets.get(s["id"])
            if w is not None and w.winfo_exists():
                cy = w.winfo_rooty() + w.winfo_height() / 2
                positions.append((cy, s))
        positions.sort(key=lambda p: p[0])

        idx = 0
        for cy, _s in positions:
            if drop_y_root > cy:
                idx += 1
            else:
                break

        if not positions:
            new_order = 0
        elif idx == 0:
            new_order = positions[0][1].get("order", 0) - 1
        elif idx >= len(positions):
            new_order = positions[-1][1].get("order", 0) + 1
        else:
            a = positions[idx - 1][1].get("order", 0)
            b = positions[idx][1].get("order", 0)
            new_order = (a + b) / 2

        now_iso = datetime.now().isoformat()
        if target_status == "in_progress" and not task.get("started_at"):
            task["started_at"] = now_iso
        if target_status == "done":
            task["completed_at"] = now_iso
        elif old_status == "done":
            task["completed_at"] = None

        task["status"] = target_status
        task["order"] = new_order

        save_data(self.data)
        self.render()

    # ---- Task detail / notes dialog -----------------------------------
    def _open_task_dialog(self, task_id: int):
        task = self._find_task(task_id)
        if task is None:
            return

        win = ctk.CTkToplevel(self.root)
        win.title("Task Details")
        win.geometry("380x480")
        win.configure(fg_color=BG)
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        # NOTE: CustomTkinter multiplies every geometry() request by the
        # detected Windows DPI scale (1.5x on this machine), so a requested
        # height of 480 actually renders at 720px — a plain "380x640" request
        # was rendering at 960px, taller than even a 800px-tall screen. The
        # content is also wrapped in a scrollable frame so it can't overflow
        # regardless of how much content (e.g. many subtasks) exists.
        body = ctk.CTkScrollableFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=0, pady=0)

        ctk.CTkLabel(
            body, text="Title", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(20, 4))
        title_entry = ctk.CTkEntry(
            body, fg_color=ENTRY_BG, border_color=BORDER, border_width=1, corner_radius=10, height=38
        )
        title_entry.insert(0, task["text"])
        title_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(
            body, text="Notes", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(16, 4))
        notes_box = ctk.CTkTextbox(
            body,
            fg_color=ENTRY_BG,
            border_color=BORDER,
            border_width=1,
            corner_radius=10,
            height=110,
            font=("Segoe UI", 12),
        )
        notes_box.insert("1.0", task.get("notes", ""))
        notes_box.pack(fill="both", padx=20, expand=False)

        ctk.CTkLabel(
            body, text="Schedule — start auto-moves this to In Progress",
            font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY, wraplength=340, justify="left",
        ).pack(anchor="w", padx=20, pady=(14, 4))
        schedule_row = ctk.CTkFrame(body, fg_color="transparent")
        schedule_row.pack(fill="x", padx=20)

        start_labels, label_to_start = build_time_options(iso_to_hhmm(task.get("start_time")))
        end_labels, label_to_end = build_time_options(iso_to_hhmm(task.get("end_time")))
        current_start_label = hhmm_display(iso_to_hhmm(task["start_time"])) if task.get("start_time") else "No time"
        current_end_label = hhmm_display(iso_to_hhmm(task["end_time"])) if task.get("end_time") else "No time"

        start_col = ctk.CTkFrame(schedule_row, fg_color="transparent")
        start_col.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ctk.CTkLabel(start_col, text="Start", font=("Segoe UI", 10), text_color=TEXT_MUTED).pack(anchor="w")
        start_menu = ctk.CTkOptionMenu(
            start_col, values=start_labels, fg_color=ENTRY_BG, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, corner_radius=10, dynamic_resizing=False,
        )
        start_menu.set(current_start_label)
        start_menu.pack(fill="x")

        end_col = ctk.CTkFrame(schedule_row, fg_color="transparent")
        end_col.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ctk.CTkLabel(end_col, text="End", font=("Segoe UI", 10), text_color=TEXT_MUTED).pack(anchor="w")
        end_menu = ctk.CTkOptionMenu(
            end_col, values=end_labels, fg_color=ENTRY_BG, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, corner_radius=10, dynamic_resizing=False,
        )
        end_menu.set(current_end_label)
        end_menu.pack(fill="x")

        schedule_error = ctk.CTkLabel(body, text="", font=("Segoe UI", 10), text_color=DANGER_TEXT)
        schedule_error.pack(anchor="w", padx=20, pady=(4, 0))

        ctk.CTkLabel(
            body, text="Priority", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(12, 4))
        priority_row = ctk.CTkFrame(body, fg_color="transparent")
        priority_row.pack(fill="x", padx=20, pady=(0, 10))

        current_priority = task.get("priority", "medium")
        priority_dot = ctk.CTkLabel(
            priority_row, text="", width=14, height=14, corner_radius=7,
            fg_color=PRIORITY_COLORS[current_priority],
        )
        priority_dot.pack(side="left", padx=(0, 8))

        label_to_priority = {v: k for k, v in PRIORITY_LABELS.items()}

        def on_priority_change(choice):
            priority_dot.configure(fg_color=PRIORITY_COLORS[label_to_priority[choice]])

        priority_menu = ctk.CTkOptionMenu(
            priority_row,
            values=[PRIORITY_LABELS[p] for p in PRIORITY_ORDER],
            fg_color=ENTRY_BG,
            button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            corner_radius=10,
            command=on_priority_change,
        )
        priority_menu.set(PRIORITY_LABELS[current_priority])
        priority_menu.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            body, text="Status", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(0, 4))
        status_menu = ctk.CTkOptionMenu(
            body,
            values=[STATUS_LABELS[s] for s in STATUS_ORDER],
            fg_color=ENTRY_BG,
            button_color=ACCENT,
            button_hover_color=ACCENT_HOVER,
            corner_radius=10,
        )
        status_menu.set(STATUS_LABELS[task["status"]])
        status_menu.pack(fill="x", padx=20, pady=(0, 10))

        parent_id = task.get("parent_id")
        parent_task = self._find_task(parent_id) if parent_id is not None else None
        if parent_task is not None:
            def open_parent():
                win.destroy()
                self.root.after(50, lambda: self._open_task_dialog(parent_id))

            parent_link = ctk.CTkLabel(
                body, text=f"↳ Subtask of: {parent_task['text'][:50]}", font=("Segoe UI", 11),
                text_color=ACCENT, anchor="w", cursor="hand2",
            )
            parent_link.pack(anchor="w", padx=20, pady=(0, 8))
            parent_link.bind("<Button-1>", lambda _e: open_parent())

        ctk.CTkLabel(
            body, text="Subtasks", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(4, 4))

        subtasks_frame = ctk.CTkScrollableFrame(body, fg_color=CARD, corner_radius=10, height=110)
        subtasks_frame.pack(fill="x", padx=20)

        def open_subtask(sub_id):
            win.destroy()
            self.root.after(50, lambda: self._open_task_dialog(sub_id))

        def render_subtasks():
            for child in subtasks_frame.winfo_children():
                child.destroy()
            children = [t for t in self.data["tasks"] if t.get("parent_id") == task_id]
            if not children:
                ctk.CTkLabel(
                    subtasks_frame, text="No subtasks yet.", font=("Segoe UI", 11), text_color=TEXT_MUTED,
                ).pack(anchor="w", padx=6, pady=6)
                return
            for sub in children:
                row = ctk.CTkFrame(subtasks_frame, fg_color="transparent")
                row.pack(fill="x", pady=2)
                ctk.CTkFrame(
                    row, width=8, height=8, corner_radius=4, fg_color=COLUMN_ACCENTS[sub["status"]],
                ).pack(side="left", padx=(4, 6))
                lbl = ctk.CTkLabel(
                    row,
                    text=sub["text"],
                    font=self.strike_font if sub["status"] == "done" else self.normal_font,
                    text_color=TEXT_MUTED if sub["status"] == "done" else TEXT_PRIMARY,
                    anchor="w", cursor="hand2",
                )
                lbl.pack(side="left", fill="x", expand=True)
                lbl.bind("<Button-1>", lambda _e, sid=sub["id"]: open_subtask(sid))

        render_subtasks()

        add_sub_row = ctk.CTkFrame(body, fg_color="transparent")
        add_sub_row.pack(fill="x", padx=20, pady=(8, 14))
        new_sub_entry = ctk.CTkEntry(
            add_sub_row, placeholder_text="Add a subtask...", fg_color=ENTRY_BG,
            border_color=BORDER, border_width=1, corner_radius=10, height=34, font=("Segoe UI", 12),
        )
        new_sub_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        def do_add_subtask():
            self.add_task(text=new_sub_entry.get(), parent_id=task_id)
            new_sub_entry.delete(0, "end")
            render_subtasks()

        new_sub_entry.bind("<Return>", lambda _e: do_add_subtask())
        self._floating_circle_button(
            add_sub_row, "+", 34, ACCENT, ACCENT_HOVER, "white", ("Segoe UI", 15, "bold"), do_add_subtask
        ).pack(side="left")

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(10, 20), side="bottom")

        def do_delete():
            win.destroy()
            self.delete_task(task_id)

        def do_save():
            new_text = title_entry.get().strip()
            if not new_text:
                return

            start_hhmm = label_to_start.get(start_menu.get())
            end_hhmm = label_to_end.get(end_menu.get())
            new_start = hhmm_to_iso(start_hhmm) if start_hhmm else None
            new_end = hhmm_to_iso(end_hhmm) if end_hhmm else None
            if new_start and new_end and new_end <= new_start:
                schedule_error.configure(text="End time must be after start time")
                return
            schedule_error.configure(text="")

            if new_end != task.get("end_time"):
                task["reminded_for_end_time"] = None
            task["start_time"] = new_start
            task["end_time"] = new_end

            task["text"] = new_text
            task["notes"] = notes_box.get("1.0", "end-1c").strip()
            task["priority"] = label_to_priority[priority_menu.get()]

            label_to_status = {v: k for k, v in STATUS_LABELS.items()}
            new_status = label_to_status[status_menu.get()]
            if new_status != task["status"]:
                now_iso = datetime.now().isoformat()
                if new_status == "in_progress" and not task.get("started_at"):
                    task["started_at"] = now_iso
                if new_status == "done":
                    task["completed_at"] = now_iso
                elif task["status"] == "done":
                    task["completed_at"] = None
                existing = [
                    t.get("order", 0)
                    for t in self.data["tasks"]
                    if t["status"] == new_status and t["id"] != task_id
                ]
                task["order"] = (max(existing) + 1) if existing else 0
                task["status"] = new_status

            save_data(self.data)
            win.destroy()
            self.render()

        ctk.CTkButton(
            btn_row,
            text="Delete",
            fg_color="transparent",
            hover_color=DANGER,
            text_color=DANGER_TEXT,
            corner_radius=8,
            width=70,
            command=do_delete,
        ).pack(side="left")
        ctk.CTkButton(
            btn_row,
            text="Cancel",
            fg_color="transparent",
            hover_color=CARD,
            text_color=TEXT_SECONDARY,
            corner_radius=8,
            width=70,
            command=win.destroy,
        ).pack(side="right", padx=(0, 8))
        ctk.CTkButton(
            btn_row,
            text="Save",
            fg_color=ACCENT,
            hover_color=ACCENT_HOVER,
            corner_radius=8,
            width=80,
            command=do_save,
        ).pack(side="right")

    # ---- Scheduling: auto-start + deadline reminders --------------------
    def _check_schedules(self):
        now = datetime.now()
        for t in self.data["tasks"]:
            if t["status"] == "todo" and t.get("start_time"):
                try:
                    start_dt = datetime.fromisoformat(t["start_time"])
                except ValueError:
                    start_dt = None
                if start_dt and now >= start_dt:
                    self._move_task(t["id"], "in_progress")
                    self._notify_task_started(t)

            if t["status"] != "done" and t.get("end_time"):
                try:
                    end_dt = datetime.fromisoformat(t["end_time"])
                except ValueError:
                    end_dt = None
                if end_dt and now >= end_dt and t.get("reminded_for_end_time") != t["end_time"]:
                    t["reminded_for_end_time"] = t["end_time"]
                    save_data(self.data)
                    self._show_deadline_notification(t["id"])

        self.root.after(20000, self._check_schedules)

    def _notify_task_started(self, task: dict):
        try:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass
        self._show_toast(f"▶ Started: {task['text']}", COLUMN_ACCENTS["in_progress"])

    def _show_toast(self, text: str, accent: str, duration_ms: int = 4000):
        """A small, non-interactive, self-dismissing notification banner."""
        toast = ctk.CTkToplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.configure(fg_color=CARD2)

        frame = ctk.CTkFrame(toast, fg_color=CARD2, corner_radius=10, border_width=1, border_color=accent)
        frame.pack()
        ctk.CTkLabel(
            frame, text=text, font=("Segoe UI", 11, "bold"), text_color=TEXT_PRIMARY,
            wraplength=280, justify="left",
        ).pack(padx=16, pady=12)

        self._position_notification(toast)
        self.active_notifications.append(toast)

        def close():
            if toast in self.active_notifications:
                self.active_notifications.remove(toast)
            if toast.winfo_exists():
                toast.destroy()

        toast.after(duration_ms, close)

    def _position_notification(self, win):
        win.update_idletasks()
        w, h = win.winfo_width(), win.winfo_height()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        x = sw - w - 24
        y = sh - h - 60 - (len(self.active_notifications) * (h + 10))
        win.geometry(f"+{x}+{y}")

    def _show_deadline_notification(self, task_id: int):
        task = self._find_task(task_id)
        if task is None:
            return
        try:
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except Exception:
            pass

        win = ctk.CTkToplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(fg_color=CARD2)

        frame = ctk.CTkFrame(win, fg_color=CARD2, corner_radius=12, border_width=2, border_color=PRIORITY_COLORS["high"])
        frame.pack()

        ctk.CTkLabel(
            frame, text="⏰ Time slot ended", font=("Segoe UI", 13, "bold"), text_color=TEXT_PRIMARY,
        ).pack(anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(
            frame, text=task["text"], font=("Segoe UI", 12), text_color=TEXT_SECONDARY,
            wraplength=260, justify="left",
        ).pack(anchor="w", padx=16)
        ctk.CTkLabel(
            frame, text=f"Was due at {fmt_12h(task.get('end_time'))}",
            font=("Segoe UI", 10), text_color=TEXT_MUTED,
        ).pack(anchor="w", padx=16, pady=(2, 10))

        def close():
            if win in self.active_notifications:
                self.active_notifications.remove(win)
            if win.winfo_exists():
                win.destroy()

        def extend(minutes):
            base = datetime.fromisoformat(task["end_time"]) if task.get("end_time") else datetime.now()
            task["end_time"] = (base + timedelta(minutes=minutes)).isoformat()
            task["reminded_for_end_time"] = None
            save_data(self.data)
            self.render()
            close()

        quick_row = ctk.CTkFrame(frame, fg_color="transparent")
        quick_row.pack(fill="x", padx=16, pady=(0, 8))
        for label, minutes in (("+15m", 15), ("+30m", 30), ("+1h", 60)):
            ctk.CTkButton(
                quick_row, text=label, width=60, height=28, corner_radius=8,
                fg_color=ENTRY_BG, hover_color=CARD, text_color=TEXT_PRIMARY,
                font=("Segoe UI", 11), command=lambda m=minutes: extend(m),
            ).pack(side="left", padx=(0, 6))

        edit_row = ctk.CTkFrame(frame, fg_color="transparent")
        edit_row.pack(fill="x", padx=16, pady=(0, 10))
        ctk.CTkLabel(
            edit_row, text="New end time:", font=("Segoe UI", 10), text_color=TEXT_MUTED,
        ).pack(side="left", padx=(0, 6))

        notif_labels, notif_label_to_hhmm = build_time_options(iso_to_hhmm(task.get("end_time")))
        time_menu = ctk.CTkOptionMenu(
            edit_row, values=notif_labels, fg_color=ENTRY_BG, button_color=ACCENT,
            button_hover_color=ACCENT_HOVER, corner_radius=8, width=110, height=28,
            font=("Segoe UI", 11), dynamic_resizing=False,
        )
        time_menu.set(notif_labels[1] if len(notif_labels) > 1 else notif_labels[0])
        time_menu.pack(side="left", padx=(0, 6))

        def apply_custom_time():
            hhmm = notif_label_to_hhmm.get(time_menu.get())
            new_iso = hhmm_to_iso(hhmm) if hhmm else None
            if new_iso is None:
                return
            task["end_time"] = new_iso
            task["reminded_for_end_time"] = None
            save_data(self.data)
            self.render()
            close()

        ctk.CTkButton(
            edit_row, text="Set", width=50, height=28, corner_radius=8,
            fg_color=ACCENT, hover_color=ACCENT_HOVER, font=("Segoe UI", 11, "bold"),
            command=apply_custom_time,
        ).pack(side="left")

        action_row = ctk.CTkFrame(frame, fg_color="transparent")
        action_row.pack(fill="x", padx=16, pady=(0, 14))

        def mark_done():
            self._move_task(task_id, "done")
            close()

        ctk.CTkButton(
            action_row, text="Mark Done", fg_color=SUCCESS, hover_color="#3fb873",
            text_color="#0a2412", corner_radius=8, height=30, command=mark_done,
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ctk.CTkButton(
            action_row, text="Dismiss", fg_color="transparent", hover_color=CARD,
            text_color=TEXT_SECONDARY, corner_radius=8, height=30, width=80, command=close,
        ).pack(side="right")

        self._position_notification(win)
        self.active_notifications.append(win)

    # ---- Notes (separate from the todo board) --------------------------
    def _find_note(self, note_id):
        for n in self.notes_data["notes"]:
            if n["id"] == note_id:
                return n
        return None

    def delete_note(self, note_id: int):
        self.notes_data["notes"] = [n for n in self.notes_data["notes"] if n["id"] != note_id]
        save_notes(self.notes_data)
        self.render_notes()

    def render_notes(self):
        if self.notes_list_frame is None:
            return
        for child in self.notes_list_frame.winfo_children():
            child.destroy()

        notes = sorted(
            self.notes_data["notes"], key=lambda n: n.get("created_at") or "", reverse=True
        )
        if not notes:
            ctk.CTkLabel(
                self.notes_list_frame,
                text="No notes yet.\nUse “+ New Note” to jot down logins, company info, or anything else.",
                font=("Segoe UI", 13),
                text_color=TEXT_MUTED,
                justify="center",
            ).pack(pady=40)
            return

        for note in notes:
            self._build_note_card(self.notes_list_frame, note)

    def _build_note_card(self, parent, note: dict):
        card = ctk.CTkFrame(parent, fg_color=CARD2, corner_radius=10)
        card.pack(fill="x", pady=5, padx=2)

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True, padx=(16, 6), pady=12)

        click_targets = [card, content]

        body = note.get("text", "")
        title_text = note.get("title") or (body.splitlines()[0][:70] if body else "Untitled note")
        title_lbl = ctk.CTkLabel(
            content, text=title_text, font=("Segoe UI", 14, "bold"), text_color=TEXT_PRIMARY,
            anchor="w", justify="left", wraplength=800,
        )
        title_lbl.pack(fill="x", anchor="w")
        click_targets.append(title_lbl)

        if body:
            preview = "\n".join(body.splitlines()[:3])
            if len(preview) > 260:
                preview = preview[:260] + "…"
            preview_lbl = ctk.CTkLabel(
                content, text=preview, font=("Segoe UI", 12), text_color=TEXT_SECONDARY,
                anchor="w", justify="left", wraplength=800,
            )
            preview_lbl.pack(fill="x", anchor="w", pady=(4, 0))
            click_targets.append(preview_lbl)

        meta = f"Added {fmt_datetime_12h(note.get('created_at'))}"
        if note.get("updated_at") and note.get("updated_at") != note.get("created_at"):
            meta += f"   ·   Edited {fmt_datetime_12h(note.get('updated_at'))}"
        meta_lbl = ctk.CTkLabel(
            content, text=meta, font=("Segoe UI", 10), text_color=TEXT_MUTED, anchor="w",
        )
        meta_lbl.pack(fill="x", anchor="w", pady=(8, 0))
        click_targets.append(meta_lbl)

        del_btn = self._floating_circle_button(
            card, "✕", 24, CARD, DANGER, TEXT_SECONDARY, ("Segoe UI", 12),
            lambda nid=note["id"]: self.delete_note(nid),
        )
        del_btn.pack(side="right", padx=(0, 10), pady=12, anchor="n")

        for widget in click_targets:
            widget.bind("<Button-1>", lambda _e, nid=note["id"]: self._open_note_dialog(nid))
            widget.configure(cursor="hand2")

        return card

    def _open_note_dialog(self, note_id):
        note = self._find_note(note_id) if note_id is not None else None

        win = ctk.CTkToplevel(self.root)
        win.title("Edit Note" if note else "New Note")
        win.geometry("460x600" if note else "460x480")
        win.configure(fg_color=BG)
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        ctk.CTkLabel(
            win, text="Title (optional)", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(20, 4))
        title_entry = ctk.CTkEntry(
            win, fg_color=ENTRY_BG, border_color=BORDER, border_width=1, corner_radius=10, height=38
        )
        title_entry.insert(0, note.get("title", "") if note else "")
        title_entry.pack(fill="x", padx=20)

        ctk.CTkLabel(
            win, text="Note", font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY
        ).pack(anchor="w", padx=20, pady=(16, 4))
        body_box = ctk.CTkTextbox(
            win, fg_color=ENTRY_BG, border_color=BORDER, border_width=1, corner_radius=10,
            height=170, font=("Segoe UI", 12),
        )
        body_box.insert("1.0", note.get("text", "") if note else "")
        body_box.pack(fill="x", padx=20)

        if note:
            meta = f"Created {fmt_datetime_12h(note.get('created_at'))}"
            if note.get("updated_at") and note.get("updated_at") != note.get("created_at"):
                meta += f"   ·   Last updated {fmt_datetime_12h(note.get('updated_at'))}"
            ctk.CTkLabel(
                win, text=meta, font=("Segoe UI", 11, "bold"), text_color=TEXT_SECONDARY,
            ).pack(anchor="w", padx=20, pady=(12, 4))

            ctk.CTkLabel(
                win, text="History", font=("Segoe UI", 11, "bold"), text_color=TEXT_MUTED,
            ).pack(anchor="w", padx=20, pady=(4, 4))
            history_frame = ctk.CTkScrollableFrame(
                win, fg_color=ENTRY_BG, corner_radius=10, height=110, scrollbar_button_color=BORDER
            )
            history_frame.pack(fill="x", padx=20)

            entries = sorted(
                note.get("history", []), key=lambda h: h.get("at") or "", reverse=True
            )
            if not entries:
                ctk.CTkLabel(
                    history_frame, text="No history recorded.", font=("Segoe UI", 10),
                    text_color=TEXT_MUTED,
                ).pack(anchor="w", padx=10, pady=6)
            else:
                icons = {"created": "✨ Created", "edited": "✏️ Edited"}
                for h in entries:
                    label = icons.get(h.get("action"), h.get("action", "Updated").capitalize())
                    ctk.CTkLabel(
                        history_frame,
                        text=f"{label} — {fmt_datetime_12h(h.get('at'))}",
                        font=("Segoe UI", 10),
                        text_color=TEXT_SECONDARY,
                        anchor="w",
                    ).pack(fill="x", anchor="w", padx=10, pady=2)

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(14, 20), side="bottom")

        def do_delete():
            win.destroy()
            self.delete_note(note_id)

        def do_save():
            title = title_entry.get().strip()
            text = body_box.get("1.0", "end-1c").strip()
            if not title and not text:
                return
            now_iso = datetime.now().isoformat()
            if note:
                note["title"] = title
                note["text"] = text
                note["updated_at"] = now_iso
                note.setdefault("history", []).append({"at": now_iso, "action": "edited"})
            else:
                new_note = {
                    "id": self.notes_data["next_id"],
                    "title": title,
                    "text": text,
                    "created_at": now_iso,
                    "updated_at": now_iso,
                    "history": [{"at": now_iso, "action": "created"}],
                }
                self.notes_data["next_id"] += 1
                self.notes_data["notes"].append(new_note)
            save_notes(self.notes_data)
            win.destroy()
            self.render_notes()

        if note:
            ctk.CTkButton(
                btn_row, text="Delete", fg_color="transparent", hover_color=DANGER,
                text_color=DANGER_TEXT, corner_radius=8, width=70, command=do_delete,
            ).pack(side="left")

        ctk.CTkButton(
            btn_row, text="Cancel", fg_color="transparent", hover_color=CARD,
            text_color=TEXT_SECONDARY, corner_radius=8, width=70, command=win.destroy,
        ).pack(side="right", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Save", fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8,
            width=80, command=do_save,
        ).pack(side="right")

    # ---- History viewer -------------------------------------------------
    def open_history(self):
        history = load_history()
        today_str = date.today().isoformat()
        today_has_tasks = bool(self.data["tasks"]) or bool(history.get(today_str))

        win = ctk.CTkToplevel(self.root)
        win.title("History")
        win.geometry("700x560")
        win.configure(fg_color=BG)
        win.transient(self.root)
        win.grab_set()

        ctk.CTkLabel(
            win, text="Past Todo Lists", font=("Segoe UI", 18, "bold"), text_color=TEXT_PRIMARY
        ).pack(anchor="w", padx=20, pady=(18, 10))

        if not history and not today_has_tasks:
            ctk.CTkLabel(
                win,
                text="No history yet — add some tasks and check back tomorrow!",
                font=("Segoe UI", 13),
                text_color=TEXT_MUTED,
            ).pack(padx=20, pady=40)
            ctk.CTkButton(
                win, text="Close", fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8,
                command=win.destroy,
            ).pack(pady=(0, 16))
            return

        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        body.grid_columnconfigure(0, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        date_list = ctk.CTkScrollableFrame(body, fg_color=CARD, corner_radius=12, width=170)
        date_list.grid(row=0, column=0, sticky="nsw", padx=(0, 12))

        detail = ctk.CTkScrollableFrame(body, fg_color="transparent")
        detail.grid(row=0, column=1, sticky="nsew")

        other_dates = [d for d in history.keys() if d != today_str]
        dates = ([today_str] if today_has_tasks else []) + sorted(other_dates, reverse=True)
        date_buttons = {}
        export_vars = {}
        current_day = {"value": dates[0] if dates else None}

        def get_day_tasks(day: str):
            if day == today_str:
                live = self.data["tasks"]
                live_ids = {t["id"] for t in live}
                already_recorded = [t for t in history.get(today_str, []) if t["id"] not in live_ids]
                return live + already_recorded
            return history.get(day, [])

        def show_date(day: str):
            current_day["value"] = day
            for child in detail.winfo_children():
                child.destroy()
            for d, btn in date_buttons.items():
                btn.configure(fg_color=ACCENT if d == day else "transparent")

            if day == today_str:
                nice_date = "Today — " + datetime.now().strftime("%A, %B %d")
            else:
                try:
                    nice_date = datetime.strptime(day, "%Y-%m-%d").strftime("%A, %B %d")
                except ValueError:
                    nice_date = day

            ctk.CTkLabel(
                detail, text=nice_date, font=("Segoe UI", 15, "bold"), text_color=TEXT_PRIMARY
            ).pack(anchor="w", pady=(0, 10))

            day_tasks = [migrate_task(dict(t)) for t in get_day_tasks(day)]
            if not day_tasks:
                ctk.CTkLabel(detail, text="No tasks recorded.", text_color=TEXT_MUTED).pack(anchor="w")
                return

            for status in STATUS_ORDER:
                subset = sorted(
                    [t for t in day_tasks if t["status"] == status], key=lambda t: t.get("order", 0)
                )
                if not subset:
                    continue

                section = ctk.CTkFrame(detail, fg_color="transparent")
                section.pack(fill="x", pady=(4, 4))
                dot = ctk.CTkFrame(
                    section, width=8, height=8, corner_radius=4, fg_color=COLUMN_ACCENTS[status]
                )
                dot.pack(side="left", padx=(0, 6))
                dot.pack_propagate(False)
                ctk.CTkLabel(
                    section, text=f"{STATUS_LABELS[status]} ({len(subset)})",
                    font=("Segoe UI", 12, "bold"), text_color=TEXT_SECONDARY,
                ).pack(side="left")

                for t in subset:
                    row = ctk.CTkFrame(detail, fg_color=CARD, corner_radius=8)
                    row.pack(fill="x", pady=3)
                    ctk.CTkLabel(
                        row,
                        text=t["text"],
                        font=self.strike_font if status == "done" else self.normal_font,
                        text_color=TEXT_MUTED if status == "done" else TEXT_PRIMARY,
                        anchor="w",
                        justify="left",
                        wraplength=400,
                    ).pack(fill="x", anchor="w", padx=12, pady=(8, 0))

                    times = f"Created {fmt_12h(t.get('created_at'))}"
                    if t.get("started_at"):
                        times += f"  ·  Started {fmt_12h(t.get('started_at'))}"
                    if t.get("completed_at"):
                        times += (
                            f"  ·  Completed {fmt_12h(t.get('completed_at'))}"
                            f"  ·  Took {fmt_duration(t.get('created_at'), t.get('completed_at'))}"
                        )
                    ctk.CTkLabel(
                        row, text=times, font=("Segoe UI", 10), text_color=TEXT_MUTED,
                        anchor="w", justify="left", wraplength=400,
                    ).pack(fill="x", anchor="w", padx=12, pady=(2, 0 if t.get("notes") else 8))

                    if t.get("notes"):
                        ctk.CTkLabel(
                            row,
                            text=t["notes"],
                            font=("Segoe UI", 11),
                            text_color=TEXT_SECONDARY,
                            anchor="w",
                            justify="left",
                            wraplength=400,
                        ).pack(fill="x", anchor="w", padx=12, pady=(4, 8))

        for d in dates:
            row_frame = ctk.CTkFrame(date_list, fg_color="transparent")
            row_frame.pack(fill="x", pady=2, padx=4)

            var = ctk.BooleanVar(value=False)
            chk = ctk.CTkCheckBox(
                row_frame, text="", variable=var, width=18, checkbox_width=18, checkbox_height=18,
                corner_radius=4, fg_color=ACCENT, hover_color=ACCENT_HOVER, border_color=BORDER,
            )
            chk.pack(side="left", padx=(2, 4))
            export_vars[d] = var

            label = "Today" if d == today_str else None
            if label is None:
                try:
                    label = datetime.strptime(d, "%Y-%m-%d").strftime("%b %d")
                except ValueError:
                    label = d
            btn = ctk.CTkButton(
                row_frame,
                text=label,
                fg_color="transparent",
                hover_color=CARD2,
                text_color=TEXT_PRIMARY,
                corner_radius=8,
                anchor="w",
                command=lambda dd=d: show_date(dd),
            )
            btn.pack(side="left", fill="x", expand=True)
            date_buttons[d] = btn

        if dates:
            show_date(dates[0])

        def do_export():
            selected = [d for d, v in export_vars.items() if v.get()]
            if not selected and current_day["value"]:
                selected = [current_day["value"]]
            if not selected:
                return
            self.export_pdf(selected, win)

        bottom_bar = ctk.CTkFrame(win, fg_color="transparent")
        bottom_bar.pack(fill="x", padx=20, pady=(0, 18))

        ctk.CTkLabel(
            bottom_bar,
            text="Tip: check one or more days to include in the report",
            font=("Segoe UI", 11),
            text_color=TEXT_MUTED,
        ).pack(side="left")

        ctk.CTkButton(
            bottom_bar, text="Close", fg_color=CARD, hover_color=CARD2, text_color=TEXT_PRIMARY,
            corner_radius=8, width=90, command=win.destroy,
        ).pack(side="right")
        ctk.CTkButton(
            bottom_bar, text="⬇ Export PDF", fg_color=ACCENT, hover_color=ACCENT_HOVER,
            text_color="white", corner_radius=8, width=130, command=do_export,
        ).pack(side="right", padx=(0, 8))

    # ---- PDF export -----------------------------------------------------
    def export_pdf(self, days: list, parent_window=None):
        history = load_history()
        today_str = date.today().isoformat()

        def get_tasks(day):
            if day == today_str:
                live = self.data["tasks"]
                live_ids = {t["id"] for t in live}
                # tasks completed & then deleted/cleared earlier today are recorded
                # into history[today] immediately (see record_completed_task) since
                # they'd otherwise vanish from the live list before day-rollover archives it
                already_recorded = [t for t in history.get(today_str, []) if t["id"] not in live_ids]
                return live + already_recorded
            return history.get(day, [])

        days_sorted = sorted(set(days))
        tasks_by_day = {d: [migrate_task(dict(t)) for t in get_tasks(d)] for d in days_sorted}

        if len(days_sorted) == 1:
            default_name = f"Todo_Report_{days_sorted[0]}.pdf"
        else:
            default_name = f"Todo_Report_{days_sorted[0]}_to_{days_sorted[-1]}.pdf"

        downloads_dir = Path.home() / "Downloads"
        initial_dir = str(downloads_dir) if downloads_dir.exists() else str(APP_DIR)

        path = filedialog.asksaveasfilename(
            parent=parent_window or self.root,
            defaultextension=".pdf",
            initialfile=default_name,
            initialdir=initial_dir,
            filetypes=[("PDF files", "*.pdf")],
            title="Save Todo Report",
        )
        if not path:
            return

        period_sections = [
            self.weekly_board.report_section(),
            self.half_year_board.report_section(),
            self.yearly_board.report_section(),
        ]

        try:
            build_pdf_report(path, days_sorted, tasks_by_day, period_sections)
        except Exception as e:
            messagebox.showerror("Export failed", f"Could not create the PDF:\n{e}", parent=parent_window or self.root)
            return

        messagebox.showinfo("Report saved", f"Saved report to:\n{path}", parent=parent_window or self.root)

    # ---- Window / tray lifecycle -------------------------------------
    def hide_window(self):
        self.root.withdraw()

    def show_window(self):
        self.data = load_data()
        self.render()
        self.weekly_board.reload()
        self.weekly_board.render()
        self.half_year_board.reload()
        self.half_year_board.render()
        self.yearly_board.reload()
        self.yearly_board.render()
        self.mail_monitor.refresh()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def quit_app(self):
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.after(0, self.root.destroy)

    def _start_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Open Daily Todo", lambda: self.action_queue.put("show"), default=True),
            pystray.MenuItem("Quit", lambda: self.action_queue.put("quit")),
        )
        self.tray_icon = pystray.Icon("daily_todo", make_tray_image(), "Daily Todo", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _poll_queue(self):
        try:
            while True:
                action = self.action_queue.get_nowait()
                if action == "show":
                    self.show_window()
                elif action == "quit":
                    self.quit_app()
                    return
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    TodoApp().run()
