import calendar
import json
import os
import shutil
import subprocess
import sys
import time
import traceback
import webbrowser
from datetime import date, datetime, timedelta

# Kivy Framework Imports
import kivy

kivy.require('2.0.0')

from kivy.app import App
from kivy.clock import Clock
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.dropdown import DropDown
from kivy.uix.spinner import Spinner
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.utils import platform

# Export Modules (fpdf2 is required: pip install fpdf2)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
from fpdf import FPDF
from fpdf.enums import MethodReturnValue, XPos, YPos

# Storage Locations
APP_DIR = os.path.dirname(os.path.abspath(__file__))

if platform == 'android':
    from android.storage import app_storage_path

    BASE_DIR = app_storage_path()
else:
    BASE_DIR = APP_DIR

NOTES_FILE = os.path.join(BASE_DIR, 'annals_notes.json')
EXCEL_FILE = os.path.join(BASE_DIR, 'annals_notes.xlsx')
DOCX_FILE = os.path.join(BASE_DIR, 'annals_notes.docx')
PDF_FILE = os.path.join(BASE_DIR, 'annals_notes.pdf')
ERROR_LOG_FILE = os.path.join(BASE_DIR, 'error_log.txt')

# Optional Unicode font (e.g. Noto Sans Bengali). Put a .ttf at fonts/unicode.ttf
# to get non-Latin text in PDF exports and in the note editor.
UNICODE_FONT_PATH = os.path.join(APP_DIR, 'fonts', 'unicode.ttf')

DEFAULT_COMMUNITIES = ['PLT', 'Dum Dum']
ALL_COMMUNITIES_LABEL = 'All communities'
STORE_VERSION = 2


def unicode_font_available():
    return os.path.exists(UNICODE_FONT_PATH)


# Core Helper Functions
def log_error(error_message, exc_info=None):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        with open(ERROR_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f'[{timestamp}] {error_message}\n')
            if exc_info:
                traceback.print_exception(*exc_info, file=f)
            f.write('-' * 60 + '\n')
    except Exception as e:
        print(f'Failed to log error to file: {e}')


def parse_note_data(val):
    """Extracts (text, community) from the legacy v1 note format."""
    if val is None:
        return '', 'PLT'
    if isinstance(val, dict):
        return str(val.get('text', '')).strip(), str(val.get('community', 'PLT')).strip() or 'PLT'
    return str(val).strip(), 'PLT'


# File Opening Helper
def open_file_with_default_app(file_path):
    """Opens/shares a generated file. Returns True if a viewer/share sheet was launched."""
    if not os.path.exists(file_path):
        return False
    try:
        if platform == 'android':
            # Needs `androidstorage4kivy` in buildozer requirements. It handles
            # FileProvider URIs, so we avoid the FileUriExposedException that
            # Uri.fromFile() causes on Android 7+.
            from androidstorage4kivy import SharedStorage, ShareSheet
            try:
                # Keep a copy in shared storage so the user can find it later.
                SharedStorage().copy_to_shared(file_path)
            except Exception:
                log_error('Copy to shared storage failed', sys.exc_info())
            ShareSheet().share_file(file_path)
            return True
        elif platform == 'win':
            os.startfile(os.path.abspath(file_path))
        elif platform == 'macosx':
            import subprocess
            subprocess.call(['open', file_path])
        else:
            webbrowser.open('file://' + os.path.abspath(file_path))
        return True
    except Exception:
        log_error(f'Failed to open generated file: {file_path}', sys.exc_info())
        return False


# File-lock helper
def _try_release_file(file_path):
    """
    Windows-only: if *file_path* is locked by another process, attempt to
    close the known viewer / editor that holds it, then verify the file is
    writable.  Returns True when the file can be written (or does not yet
    exist).  Always returns True on non-Windows platforms.
    """
    if not os.path.exists(file_path):
        return True          # nothing to release
    if platform != 'win':
        return True          # locking only matters on Windows

    def _writable():
        try:
            with open(file_path, 'r+b'):
                return True
        except PermissionError:
            return False

    if _writable():
        return True          # file is not locked

    # Map each extension to the process names of apps that lock it
    ext = os.path.splitext(file_path)[1].lower()
    viewer_procs = {
        '.pdf':  ['AcroRd32.exe', 'Acrobat.exe', 'SumatraPDF.exe',
                  'FoxitReader.exe', 'FoxitPDFReader.exe'],
        '.xlsx': ['EXCEL.EXE'],
        '.docx': ['WINWORD.EXE'],
    }
    for proc_name in viewer_procs.get(ext, []):
        try:
            subprocess.run(
                ['taskkill', '/f', '/im', proc_name],
                capture_output=True, timeout=5
            )
        except Exception:
            pass

    time.sleep(0.5)          # give Windows time to release handles
    return _writable()


# Data Store
class NoteStore:
    """
    On-disk format (v2):
    {
      "version": 2,
      "communities": ["PLT", "Dum Dum"],
      "selected": "PLT",
      "notes": {"2026-10-03": {"PLT": "text", "Dum Dum": "text"}}
    }
    Legacy v1 files (date -> text or {text, community}) are migrated automatically
    and the original is kept as <file>.v1.bak.
    """

    def __init__(self, path):
        self.path = path
        self.communities = list(DEFAULT_COMMUNITIES)
        self.selected = self.communities[0]
        self.notes = {}
        self._load()

    # --- loading / saving ---
    def _load(self):
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            log_error('Error loading notes file', sys.exc_info())
            self._backup(self.path + '.corrupt-' + datetime.now().strftime('%Y%m%d-%H%M%S'))
            return

        if isinstance(data, dict) and data.get('version') == STORE_VERSION:
            self._load_v2(data)
        else:
            self._backup(self.path + '.v1.bak')
            self._migrate_v1(data)
            self.save()

    def _backup(self, dest):
        try:
            shutil.copy2(self.path, dest)
        except Exception:
            log_error('Could not back up notes file', sys.exc_info())

    def _load_v2(self, data):
        comms = []
        for c in data.get('communities', []):
            c = str(c).strip()
            if c and c not in comms:
                comms.append(c)

        notes = {}
        raw_notes = data.get('notes', {})
        if isinstance(raw_notes, dict):
            for date_key, per_comm in raw_notes.items():
                if not isinstance(per_comm, dict):
                    continue
                for comm, text in per_comm.items():
                    comm, text = str(comm).strip(), str(text).strip()
                    if comm and text:
                        notes.setdefault(date_key, {})[comm] = text
                        if comm not in comms:
                            comms.append(comm)

        self.notes = notes
        self.communities = comms or list(DEFAULT_COMMUNITIES)
        selected = str(data.get('selected', '')).strip()
        self.selected = selected if selected in self.communities else self.communities[0]

    def _migrate_v1(self, data):
        comms = list(DEFAULT_COMMUNITIES)
        notes = {}
        if isinstance(data, dict):
            for date_key, val in data.items():
                text, comm = parse_note_data(val)
                if text:
                    notes.setdefault(date_key, {})[comm] = text
                    if comm not in comms:
                        comms.append(comm)
        self.notes = notes
        self.communities = comms
        self.selected = comms[0]

    def save(self):
        tmp_path = self.path + '.tmp'
        try:
            payload = {
                'version': STORE_VERSION,
                'communities': self.communities,
                'selected': self.selected,
                'notes': self.notes,
            }
            with open(tmp_path, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, self.path)  # atomic: never leaves a half-written file
            return True
        except Exception:
            log_error('Error saving notes file', sys.exc_info())
            return False

    # --- notes ---
    def get(self, d, community):
        return self.notes.get(d.isoformat(), {}).get(community, '')

    def has(self, d, community):
        return bool(self.get(d, community))

    def set(self, d, community, text):
        text = text.strip()
        if not text:
            return self.delete(d, community)
        self.notes.setdefault(d.isoformat(), {})[community] = text
        return self.save()

    def delete(self, d, community):
        key = d.isoformat()
        per_comm = self.notes.get(key)
        if per_comm and community in per_comm:
            del per_comm[community]
            if not per_comm:
                del self.notes[key]
            return self.save()
        return True

    def query(self, start_d=None, end_d=None, community=None):
        rows = []
        for date_str, per_comm in self.notes.items():
            try:
                d_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
            except ValueError:
                log_error(f'Bad date key in notes: {date_str}')
                continue
            if start_d and d_obj < start_d:
                continue
            if end_d and d_obj > end_d:
                continue
            for comm, text in per_comm.items():
                if community and comm != community:
                    continue
                if text:
                    rows.append((date_str, comm, text))
        rows.sort(key=lambda r: (r[0], r[1].lower()))
        return rows

    # --- communities ---
    def find_community(self, name):
        lowered = name.strip().lower()
        for c in self.communities:
            if c.lower() == lowered:
                return c
        return None

    def add_community(self, name):
        """Adds a community (or returns the existing one with the same name)."""
        name = name.strip()
        existing = self.find_community(name)
        if existing:
            return existing
        self.communities.append(name)
        self.save()
        return name

    def rename_community(self, old, new):
        new = new.strip()
        clash = self.find_community(new)
        if not new or old not in self.communities:
            return False
        if clash and clash != old:
            return False  # name already used by another community
        self.communities[self.communities.index(old)] = new
        for per_comm in self.notes.values():
            if old in per_comm:
                per_comm[new] = per_comm.pop(old)
        if self.selected == old:
            self.selected = new
        return self.save()

    def set_selected(self, name):
        if name in self.communities and name != self.selected:
            self.selected = name
            self.save()


# PDF
class PDFReport(FPDF):
    FONT = 'AnnalsUnicode'

    def __init__(self, start_date_str, end_date_str, community_label):
        super().__init__()
        self.start_date_str = start_date_str
        self.end_date_str = end_date_str
        self.community_label = community_label
        self.use_unicode = unicode_font_available()
        if self.use_unicode:
            self.add_font(self.FONT, '', fname=UNICODE_FONT_PATH)
            try:
                # Proper shaping for Bengali etc. (needs: pip install uharfbuzz)
                self.set_text_shaping(True)
            except Exception:
                log_error('Text shaping unavailable', sys.exc_info())

    def body_font(self, style='', size=10):
        if self.use_unicode:
            self.set_font(self.FONT, '', size)  # one weight only
        else:
            self.set_font('Helvetica', style, size)

    def safe(self, text):
        if self.use_unicode:
            return text
        return text.encode('latin-1', 'replace').decode('latin-1')

    def header(self):
        self.body_font('B', 16)
        self.set_text_color(44, 62, 80)
        self.cell(0, 8, 'Annals', new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L')
        self.body_font('I', 10)
        self.set_text_color(127, 140, 141)
        subtitle = f'Date Range: {self.start_date_str} to {self.end_date_str}  |  {self.community_label}'
        self.cell(0, 6, self.safe(subtitle), new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L')
        self.ln(4)
        self.set_text_color(0, 0, 0)

    def footer(self):
        self.set_y(-15)
        self.body_font('I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')
        self.set_text_color(0, 0, 0)

    def table_header(self, widths, labels=('Date', 'Community', 'Note')):
        self.body_font('B', 10)
        x, y = self.l_margin, self.get_y()
        for w, label in zip(widths, labels):
            self.rect(x, y, w, 7)
            self.set_xy(x, y)
            self.cell(w, 7, label)
            x += w
        self.set_xy(self.l_margin, y + 7)
        self.body_font('', 10)

    def wrapped_lines(self, width, text, line_height):
        lines = self.multi_cell(width, line_height, text, dry_run=True,
                                output=MethodReturnValue.LINES)
        return list(lines) or ['']

    def draw_row(self, widths, line_height, cells):
        """cells: one list of pre-wrapped lines per column; all borders share one height."""
        height = max(len(c) for c in cells) * line_height
        x0, y0 = self.l_margin, self.get_y()
        x = x0
        for w, lines in zip(widths, cells):
            self.rect(x, y0, w, height)
            self.set_xy(x, y0)
            self.multi_cell(w, line_height, '\n'.join(lines), border=0,
                            new_x=XPos.RIGHT, new_y=YPos.TOP)
            x += w
        self.set_xy(x0, y0 + height)


# Dialog & Message UI Popups
class AlertPopup(Popup):
    def __init__(self, title, text, **kwargs):
        super().__init__(**kwargs)
        self.title = title
        self.size_hint = (0.8, 0.4)

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        layout.add_widget(Label(text=text, halign='center', valign='middle'))

        ok_btn = Button(text='OK', size_hint_y=0.3)
        ok_btn.bind(on_release=self.dismiss)
        layout.add_widget(ok_btn)

        self.content = layout


class ConfirmPopup(Popup):
    def __init__(self, title, text, on_yes, **kwargs):
        super().__init__(**kwargs)
        self.title = title
        self.on_yes = on_yes
        self.size_hint = (0.8, 0.4)

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        layout.add_widget(Label(text=text, halign='center', valign='middle'))

        btn_box = BoxLayout(size_hint_y=0.3, spacing=10)
        no_btn = Button(text='Cancel')
        no_btn.bind(on_release=self.dismiss)
        yes_btn = Button(text='Delete')
        yes_btn.bind(on_release=self._confirm)
        btn_box.add_widget(no_btn)
        btn_box.add_widget(yes_btn)
        layout.add_widget(btn_box)

        self.content = layout

    def _confirm(self, instance):
        self.dismiss()
        self.on_yes()


class AddCommunityPopup(Popup):
    def __init__(self, callback, **kwargs):
        super().__init__(**kwargs)
        self.title = 'Add New Community'
        self.callback = callback
        self.size_hint = (0.8, 0.4)

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        self.input = TextInput(hint_text='Enter Community Name', multiline=False)
        layout.add_widget(self.input)

        btn_box = BoxLayout(size_hint_y=0.4, spacing=10)
        cancel_btn = Button(text='Cancel')
        cancel_btn.bind(on_release=self.dismiss)
        add_btn = Button(text='Add')
        add_btn.bind(on_release=self.submit)

        btn_box.add_widget(cancel_btn)
        btn_box.add_widget(add_btn)
        layout.add_widget(btn_box)

        self.content = layout

    def submit(self, instance):
        name = self.input.text.strip()
        self.dismiss()
        if name:
            self.callback(name)


class EditCommunityPopup(Popup):
    def __init__(self, current_name, callback, **kwargs):
        super().__init__(**kwargs)
        self.title = f'Edit Community Name ({current_name})'
        self.current_name = current_name
        self.callback = callback
        self.size_hint = (0.8, 0.4)

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        self.input = TextInput(text=current_name, multiline=False)
        layout.add_widget(self.input)

        btn_box = BoxLayout(size_hint_y=0.4, spacing=10)
        cancel_btn = Button(text='Cancel')
        cancel_btn.bind(on_release=self.dismiss)
        save_btn = Button(text='Update')
        save_btn.bind(on_release=self.submit)

        btn_box.add_widget(cancel_btn)
        btn_box.add_widget(save_btn)
        layout.add_widget(btn_box)

        self.content = layout

    def submit(self, instance):
        new_name = self.input.text.strip()
        self.dismiss()
        if new_name and new_name != self.current_name:
            self.callback(self.current_name, new_name)


class DatePickerRow(BoxLayout):
    """Day / Month / Year spinners, so nobody has to type YYYY-MM-DD on a phone."""

    MONTHS = list(calendar.month_abbr)[1:]

    def __init__(self, initial, **kwargs):
        super().__init__(spacing=4, **kwargs)
        years = [str(y) for y in range(date.today().year - 10, date.today().year + 3)]
        if str(initial.year) not in years:
            years.append(str(initial.year))
            years.sort()

        self.day = Spinner(text=str(initial.day), values=[str(i) for i in range(1, 32)],
                           size_hint_x=0.25)
        self.month = Spinner(text=self.MONTHS[initial.month - 1], values=self.MONTHS,
                             size_hint_x=0.35)
        self.year = Spinner(text=str(initial.year), values=years, size_hint_x=0.4)
        for w in (self.day, self.month, self.year):
            self.add_widget(w)

    def get_date(self):
        year = int(self.year.text)
        month = self.MONTHS.index(self.month.text) + 1
        day = min(int(self.day.text), calendar.monthrange(year, month)[1])  # clamp 31 Feb etc.
        return date(year, month, day)


class DateRangePopup(Popup):
    def __init__(self, title, communities, default_community, callback, **kwargs):
        super().__init__(**kwargs)
        self.title = title
        self.callback = callback
        self.size_hint = (0.92, 0.6)

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        grid = GridLayout(cols=2, row_default_height=44, row_force_default=True, spacing=5)
        grid.add_widget(Label(text='Start Date:', size_hint_x=0.25))
        self.start_picker = DatePickerRow(date.today().replace(day=1))
        grid.add_widget(self.start_picker)

        grid.add_widget(Label(text='End Date:', size_hint_x=0.25))
        self.end_picker = DatePickerRow(date.today())
        grid.add_widget(self.end_picker)

        grid.add_widget(Label(text='Community:', size_hint_x=0.25))
        self.comm_spinner = Spinner(
            text=default_community if default_community in communities else ALL_COMMUNITIES_LABEL,
            values=[ALL_COMMUNITIES_LABEL] + list(communities),
        )
        grid.add_widget(self.comm_spinner)
        layout.add_widget(grid)

        btn_box = BoxLayout(size_hint_y=0.3, spacing=10)
        cancel_btn = Button(text='Cancel')
        cancel_btn.bind(on_release=self.dismiss)
        confirm_btn = Button(text='Export')
        confirm_btn.bind(on_release=self.validate_and_submit)

        btn_box.add_widget(cancel_btn)
        btn_box.add_widget(confirm_btn)
        layout.add_widget(btn_box)

        self.content = layout

    def validate_and_submit(self, instance):
        start_d = self.start_picker.get_date()
        end_d = self.end_picker.get_date()
        if start_d > end_d:
            AlertPopup('Date Error', 'Start Date cannot be after End Date.').open()
            return
        community = None if self.comm_spinner.text == ALL_COMMUNITIES_LABEL else self.comm_spinner.text
        self.dismiss()
        self.callback(start_d, end_d, community)


class NoteEditorPopup(Popup):
    def __init__(self, current_date, note_text, app_instance, **kwargs):
        super().__init__(**kwargs)
        self.current_date = current_date
        self.app_instance = app_instance
        self.title = f'Notes - {current_date.strftime("%d-%m-%Y")}'
        self.size_hint = (0.95, 0.9)
        self.auto_dismiss = False  # only Close/Prev/Next/Delete leave the editor, so nothing is lost

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        # Top Control Bar
        top_bar = BoxLayout(size_hint_y=0.1, spacing=5)
        btn_del = Button(text='Delete')
        btn_del.bind(on_release=self.confirm_delete)
        btn_prev = Button(text='< Prev')
        btn_prev.bind(on_release=lambda x: self.change_date(-1))
        btn_next = Button(text='Next >')
        btn_next.bind(on_release=lambda x: self.change_date(1))
        btn_close = Button(text='Close')
        btn_close.bind(on_release=self.close_and_save)

        top_bar.add_widget(btn_del)
        top_bar.add_widget(btn_prev)
        top_bar.add_widget(btn_next)
        top_bar.add_widget(btn_close)
        layout.add_widget(top_bar)

        # Date & Community Header Display
        layout.add_widget(
            Label(
                text=f"{current_date.strftime('%a, %d %b %Y')} ({self.app_instance.selected_community})",
                font_size='18sp',
                bold=True,
                size_hint_y=0.08,
            )
        )

        # Multi-line Text Area
        text_kwargs = {}
        if unicode_font_available():
            text_kwargs['font_name'] = UNICODE_FONT_PATH
        self.text_input = TextInput(
            text=note_text, multiline=True, font_size='16sp', size_hint_y=0.82, **text_kwargs
        )
        layout.add_widget(self.text_input)

        self.content = layout

    def close_and_save(self, *args):
        self.app_instance.save_note(self.current_date, self.text_input.text)
        self.dismiss()

    def confirm_delete(self, instance):
        if not self.text_input.text.strip():
            self.dismiss()
            return
        ConfirmPopup(
            'Delete Note',
            f'Delete this note for {self.current_date.strftime("%d-%m-%Y")}\n'
            f'({self.app_instance.selected_community})?',
            on_yes=self._do_delete,
        ).open()

    def _do_delete(self):
        self.app_instance.delete_note(self.current_date)
        self.dismiss()

    def change_date(self, direction):
        self.app_instance.save_note(self.current_date, self.text_input.text)
        self.dismiss()
        new_date = self.current_date + timedelta(days=direction)
        self.app_instance.open_notes(new_date)


# Main App Root Layout
class AnnalsNotesLayout(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.padding = 5
        self.spacing = 5

        self.store = NoteStore(NOTES_FILE)
        self.current_year = date.today().year
        self.current_month = date.today().month
        self._last_today = date.today()
        self._updating_spinner = False

        # Month Title Header
        self.month_label = Label(
            text='', font_size='20sp', bold=True, size_hint_y=0.08
        )
        self.add_widget(self.month_label)

        # Navigation Bar
        nav_bar = BoxLayout(size_hint_y=0.08, spacing=3)
        btn_prev = Button(text='< Prev')
        btn_prev.bind(on_release=lambda x: self.previous_month())
        btn_today = Button(text='Today')
        btn_today.bind(on_release=lambda x: self.go_today())
        btn_next = Button(text='Next >')
        btn_next.bind(on_release=lambda x: self.next_month())

        # Export Dropdown
        self.export_dropdown = DropDown()
        for label, kind in (('Export Excel', 'excel'), ('Export Word', 'docx'), ('Export PDF', 'pdf')):
            b = Button(text=label, size_hint_y=None, height=44)
            b.bind(on_release=lambda btn, k=kind: self._handle_export_select(k))
            self.export_dropdown.add_widget(b)

        self.export_btn = Button(text='Export Notes')
        self.export_btn.bind(on_release=self.export_dropdown.open)

        nav_bar.add_widget(btn_prev)
        nav_bar.add_widget(btn_today)
        nav_bar.add_widget(btn_next)
        nav_bar.add_widget(self.export_btn)
        self.add_widget(nav_bar)

        # Community Dropdown Selector
        comm_bar = BoxLayout(size_hint_y=0.07, spacing=5)
        comm_bar.add_widget(Label(text='Community:', size_hint_x=0.3, bold=True))

        self.comm_spinner = Spinner(
            text=self.selected_community,
            values=self._get_spinner_options(),
            size_hint_x=0.7
        )
        self.comm_spinner.bind(text=self.on_community_change)
        comm_bar.add_widget(self.comm_spinner)
        self.add_widget(comm_bar)

        # Calendar View Container
        self.calendar_grid = GridLayout(cols=7, spacing=2, size_hint_y=0.77)
        self.add_widget(self.calendar_grid)

        self.create_calendar()

        # Refresh the "today" highlight if the app stays open past midnight
        Clock.schedule_interval(self._check_day_rollover, 60)

    # --- helpers ---
    @property
    def selected_community(self):
        return self.store.selected

    def _persist_failed_alert(self):
        AlertPopup('Save Error', 'Could not save your changes.\nSee error_log.txt for details.').open()

    def _get_spinner_options(self):
        return list(self.store.communities) + ['Edit Current Community', '+ Add New']

    def _set_spinner(self, text):
        """Update the spinner without triggering on_community_change."""
        self._updating_spinner = True
        try:
            self.comm_spinner.values = self._get_spinner_options()
            self.comm_spinner.text = text
        finally:
            self._updating_spinner = False

    def _check_day_rollover(self, dt):
        if date.today() != self._last_today:
            self._last_today = date.today()
            self.create_calendar()

    # --- communities ---
    def on_community_change(self, spinner, text):
        if self._updating_spinner:
            return
        if text == '+ Add New':
            self._set_spinner(self.selected_community)  # also covers Cancel
            AddCommunityPopup(callback=self.add_new_community_callback).open()
        elif text == 'Edit Current Community':
            self._set_spinner(self.selected_community)
            EditCommunityPopup(
                current_name=self.selected_community,
                callback=self.edit_community_callback
            ).open()
        else:
            self.store.set_selected(text)
            self.create_calendar()

    def add_new_community_callback(self, new_name):
        name = self.store.add_community(new_name)  # returns existing one if duplicate
        self.store.set_selected(name)
        self._set_spinner(name)
        self.create_calendar()

    def edit_community_callback(self, old_name, new_name):
        if not self.store.rename_community(old_name, new_name):
            AlertPopup('Rename Failed',
                       f'A community named "{new_name}" already exists\nor the name is invalid.').open()
            return
        self._set_spinner(self.selected_community)
        self.create_calendar()

    # --- calendar ---
    def create_calendar(self):
        self.calendar_grid.clear_widgets()
        self.month_label.text = f'{calendar.month_name[self.current_month]} {self.current_year}'

        for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']:
            self.calendar_grid.add_widget(
                Label(text=day, bold=True, size_hint_y=None, height=30)
            )

        today = date.today()
        community = self.selected_community
        for week in calendar.monthcalendar(self.current_year, self.current_month):
            for day_num in week:
                if not day_num:
                    self.calendar_grid.add_widget(Label(text=''))
                    continue

                d = date(self.current_year, self.current_month, day_num)
                has_note = self.store.has(d, community)

                btn_text = f'{day_num}\n\u2022' if has_note else str(day_num)
                btn = Button(text=btn_text, font_size='14sp')

                if d == today:
                    btn.background_color = (0.2, 0.6, 1, 1)
                elif has_note:
                    btn.background_color = (0.4, 0.8, 0.5, 1)

                btn.bind(on_release=lambda instance, dt=d: self.open_notes(dt))
                self.calendar_grid.add_widget(btn)

    def previous_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.create_calendar()

    def next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.create_calendar()

    def go_today(self):
        today = date.today()
        self.current_year = today.year
        self.current_month = today.month
        self.create_calendar()

    # --- notes ---
    def open_notes(self, d):
        text = self.store.get(d, self.selected_community)
        popup = NoteEditorPopup(current_date=d, note_text=text, app_instance=self)
        popup.open()
        Clock.schedule_once(lambda dt: setattr(popup.text_input, 'focus', True), 0.2)

    def save_note(self, d, text_content):
        if not self.store.set(d, self.selected_community, text_content):
            self._persist_failed_alert()
        self.create_calendar()

    def delete_note(self, d):
        if not self.store.delete(d, self.selected_community):
            self._persist_failed_alert()
        self.create_calendar()

    # --- exports ---
    def _deliver(self, path):
        if not open_file_with_default_app(path):
            AlertPopup('File Saved', f'Saved to:\n{path}').open()

    def _handle_export_select(self, export_type):
        self.export_dropdown.dismiss()
        runner = {
            'excel': (self._run_excel_export, 'Select Date Range for Excel'),
            'docx': (self._run_docx_export, 'Select Date Range for Word'),
            'pdf': (self._run_pdf_export, 'Select Date Range for PDF'),
        }[export_type]
        DateRangePopup(runner[1], self.store.communities, self.selected_community, runner[0]).open()

    @staticmethod
    def _range_text(start_d, end_d, community):
        return f'{start_d} to {end_d}' + (f' ({community})' if community else '')

    def _run_excel_export(self, start_d, end_d, community):
        rows = self.store.query(start_d, end_d, community)
        if not rows:
            AlertPopup('Export Excel',
                       f'No notes found for {self._range_text(start_d, end_d, community)}.').open()
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = 'Annals Notes'
            single = community is not None
            if single:
                # Community name goes in a header row; no Community column needed
                comm_cell = ws.cell(row=1, column=1, value=f'Community: {community}')
                comm_cell.font = Font(bold=True)
                ws.append(['Date', 'Note'])
                for cell in ws[2]:
                    cell.font = Font(bold=True)
                for date_str, _comm, note in rows:
                    ws.append([date_str, note])
                ws.column_dimensions['A'].width = 14
                ws.column_dimensions['B'].width = 98
                data_start_row = 3
            else:
                ws.append(['Date', 'Community', 'Note'])
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                for date_str, comm, note in rows:
                    ws.append([date_str, comm, note])
                ws.column_dimensions['A'].width = 14
                ws.column_dimensions['B'].width = 18
                ws.column_dimensions['C'].width = 80
                data_start_row = 2
            for row in ws.iter_rows(min_row=data_start_row):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')

            if not _try_release_file(EXCEL_FILE):
                AlertPopup('File In Use',
                           'annals_notes.xlsx is open in Excel.\n'
                           'Please close it and export again.').open()
                return
            wb.save(EXCEL_FILE)
            self._deliver(EXCEL_FILE)
        except Exception as e:
            log_error('Excel Export Failed', sys.exc_info())
            AlertPopup('Export Error', f'Failed to save Excel file:\n{e}').open()

    def _run_docx_export(self, start_d, end_d, community):
        rows = self.store.query(start_d, end_d, community)
        if not rows:
            AlertPopup('Export Word',
                       f'No notes found for {self._range_text(start_d, end_d, community)}.').open()
            return
        try:
            doc = Document()
            for section in doc.sections:
                section.top_margin = Inches(0.75)
                section.bottom_margin = Inches(0.75)
                section.left_margin = Inches(0.75)
                section.right_margin = Inches(0.75)

            title_p = doc.add_paragraph()
            title_run = title_p.add_run('Annals')
            title_run.font.size = Pt(18)
            title_run.font.bold = True

            sub_p = doc.add_paragraph()
            sub_run = sub_p.add_run(
                f'Date Range: {start_d} to {end_d}  |  {community or ALL_COMMUNITIES_LABEL}')
            sub_run.font.size = Pt(10)
            sub_run.font.italic = True

            single = community is not None
            if single:
                table = doc.add_table(rows=1, cols=2)
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.LEFT
                for cell, label in zip(table.rows[0].cells, ('Date', 'Note')):
                    cell.text = ''
                    cell.paragraphs[0].add_run(label).bold = True
                for date_str, _comm, note in rows:
                    row_cells = table.add_row().cells
                    row_cells[0].text = date_str
                    row_cells[1].text = note
            else:
                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.LEFT
                for cell, label in zip(table.rows[0].cells, ('Date', 'Community', 'Note')):
                    cell.text = ''
                    cell.paragraphs[0].add_run(label).bold = True
                for date_str, comm, note in rows:
                    row_cells = table.add_row().cells
                    row_cells[0].text = date_str
                    row_cells[1].text = comm
                    row_cells[2].text = note

            if not _try_release_file(DOCX_FILE):
                AlertPopup('File In Use',
                           'annals_notes.docx is open in Word.\n'
                           'Please close it and export again.').open()
                return
            doc.save(DOCX_FILE)
            self._deliver(DOCX_FILE)
        except Exception as e:
            log_error('Word Export Failed', sys.exc_info())
            AlertPopup('Export Error', f'Failed to save Word file:\n{e}').open()

    def _run_pdf_export(self, start_d, end_d, community):
        rows = self.store.query(start_d, end_d, community)
        if not rows:
            AlertPopup('Export PDF',
                       f'No notes found for {self._range_text(start_d, end_d, community)}.').open()
            return
        try:
            pdf = PDFReport(start_d.strftime('%d %b %Y'), end_d.strftime('%d %b %Y'),
                            community or ALL_COMMUNITIES_LABEL)
            pdf.alias_nb_pages()
            pdf.set_auto_page_break(auto=True, margin=15)  # set before the first page
            pdf.add_page()

            single = community is not None
            if single:
                line_height = 6
                widths = [35, 155]           # Date | Note (wider note column)
                col_labels = ('Date', 'Note')
            else:
                line_height = 6
                widths = [30, 35, 125]       # Date | Community | Note
                col_labels = ('Date', 'Community', 'Note')
            pdf.table_header(widths, col_labels)

            for date_str, comm, raw_note in rows:
                date_lines = pdf.wrapped_lines(widths[0], date_str, line_height)
                note_lines = pdf.wrapped_lines(widths[-1], pdf.safe(raw_note), line_height)

                if single:
                    row_cols = [date_lines]
                else:
                    comm_lines = pdf.wrapped_lines(widths[1], pdf.safe(comm), line_height)
                    row_cols = [date_lines, comm_lines]

                first = True
                idx = 0
                while True:
                    room = int((pdf.page_break_trigger - pdf.get_y()) / line_height + 1e-6)
                    min_lines = max(len(c) for c in row_cols) if first else 1
                    if room < max(min_lines, 1):
                        pdf.add_page()
                        pdf.table_header(widths, col_labels)
                        continue

                    segment = note_lines[idx:idx + room]
                    if first:
                        pdf.draw_row(widths, line_height, row_cols + [segment])
                    else:
                        pdf.draw_row(widths, line_height,
                                     [[''] for _ in row_cols] + [segment])
                    idx += len(segment)
                    first = False
                    if idx >= len(note_lines):
                        break

            if not _try_release_file(PDF_FILE):
                AlertPopup('File In Use',
                           'annals_notes.pdf is open in a PDF viewer.\n'
                           'Please close it and export again.').open()
                return
            pdf.output(PDF_FILE)
            self._deliver(PDF_FILE)
        except Exception as e:
            log_error('PDF Export Failed', sys.exc_info())
            AlertPopup('Export Error', f'Failed to save PDF file:\n{e}').open()


class MainApp(App):
    def build(self):
        self.title = 'Annals'
        return AnnalsNotesLayout()


if __name__ == '__main__':
    MainApp().run()
