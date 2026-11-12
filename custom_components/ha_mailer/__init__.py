"""HA_mailer - raporty Excel z migawką encji wysyłane mailem z Gmaila wg cron."""
from __future__ import annotations

import asyncio
import hashlib
import json
import logging
import math
import os
import re
import smtplib
import ssl
import stat
import uuid
from collections.abc import Callable
from datetime import datetime, timedelta
from email.message import EmailMessage
from functools import partial
from io import BytesIO

import voluptuous as vol
from croniter import croniter
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

from homeassistant.components import persistent_notification
from homeassistant.config_entries import SOURCE_IMPORT, ConfigEntry
from homeassistant.const import CONF_NAME
from homeassistant.core import HomeAssistant, ServiceCall, callback
from homeassistant.exceptions import HomeAssistantError, ServiceValidationError
import homeassistant.helpers.config_validation as cv
from homeassistant.helpers import device_registry as dr
from homeassistant.helpers import entity_registry as er
from homeassistant.helpers import issue_registry as ir
from homeassistant.helpers.device_registry import DeviceInfo
from homeassistant.helpers.dispatcher import async_dispatcher_send
from homeassistant.helpers.event import async_track_point_in_time, async_track_time_interval
from homeassistant.helpers.storage import Store
from homeassistant.util import dt as dt_util, slugify

_LOGGER = logging.getLogger(__name__)

DOMAIN = "ha_mailer"
PLATFORMS = ["sensor", "switch"]

CONF_ID = "id"
CONF_ENTITIES = "entities"  # tylko w kroku początkowym UI (zamieniane na wiersze)
CONF_CRON = "cron"
CONF_GMAIL_USER = "gmail_user"
CONF_GMAIL_PASSWORD = "gmail_app_password"
CONF_RECIPIENT = "recipient"
CONF_CC = "cc"
CONF_BCC = "bcc"
CONF_SUBJECT = "subject"
CONF_BODY = "body"
CONF_TEMPLATE = "template"
CONF_COLUMNS = "columns"
CONF_ROWS = "rows"
CONF_REPORTS = "reports"
CONF_RETRY = "retry_minutes"
CONF_RETENTION = "retention_days"
CONF_STORAGE_PATH = "storage_path"
CONF_YAML_HASH = "yaml_hash"

ROW_ENTITY = "entity"
ROW_DESC = "description"
ROW_ATTR = "attribute"
ROW_SHEET = "sheet"

# klucz kolumny -> nagłówek w Excelu (kolejność = domyślna kolejność)
COLUMN_LABELS = {
    "lp": "Lp.",
    "opis": "Opis",
    "encja": "Encja",
    "atrybut": "Atrybut",
    "nazwa": "Nazwa",
    "wartosc": "Wartość",
    "jednostka": "Jednostka",
    "zmiana": "Ostatnia zmiana",
}
DEFAULT_COLUMNS = ["lp", "opis", "encja", "nazwa", "wartosc", "jednostka", "zmiana"]
LEGACY_COLUMNS = ["lp", "encja", "nazwa", "wartosc", "jednostka", "zmiana"]  # układ z wersji 1.1

DEFAULT_CRON = "0 8 * * *"
DEFAULT_SUBJECT = "HA_mailer - raport"
DEFAULT_RETRY = 10
DEFAULT_RETENTION = 30
DEFAULT_FOLDER = "HA_mailer_reports"  # względem katalogu z configuration.yaml
DEFAULT_SHEET = "Migawka"
GMAIL_HOST = "smtp.gmail.com"
GMAIL_PORT = 465

SIGNAL = DOMAIN + "_update_{}"
DT_FMT = "%Y-%m-%d %H:%M:%S"
XLSX_MIME = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
FIRST_DATA_ROW = 5  # wiersz 1: tytuł, 2: czas migawki, 4: nagłówki, dane od wiersza 5
WARN_STATES = {"unavailable", "unknown"}
TOKEN_RE = re.compile(r"\{\{\s*([^{}]*?)\s*\}\}")
ENTITY_TOKEN_RE = re.compile(
    r"^([a-z0-9_]+\.[a-z0-9_]+)(?:\.([A-Za-z0-9_]+))?(?:#(nazwa|jednostka|zmiana))?$"
)
SPECIAL_TOKENS = {"raport", "data", "czas", "data_czas", "ostrzezenia"}
FILE_RE = re.compile(r"^HA_mailer_[a-z0-9_]+_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsx$")

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
ERR_OPIS = "Wiersze bez encji wymagają kolumny 'opis' w układzie."

DEFAULT_STATS = {
    "ok": 0,
    "fail": 0,  # liczba nieudanych PRÓB (ponowienia liczone osobno)
    "retry_ok": 0,  # wysyłki udane dopiero po ponowieniu
    "status": None,  # None | "ok" | "error"
    "last_ok": None,
    "last_attempt": None,
    "last_error": None,
    "last_error_time": None,
    "last_file_kb": None,
    "last_entities": None,  # liczba encji w ostatnim pliku (ważne przy szablonie)
    "last_warnings": None,  # liczba encji niedostępnych lub brakujących w ostatnim pliku
}


# ---------------------------------------------------------------- walidacja ---
def split_recipients(text: str) -> list[str]:
    return [r.strip() for r in str(text or "").replace(";", ",").split(",") if r.strip()]


def _cron(value):
    value = " ".join(str(value).split())
    if not croniter.is_valid(value):
        raise vol.Invalid("Nieprawidłowe wyrażenie cron.")
    return value


def _email(value):
    value = str(value).strip()
    if not EMAIL_RE.match(value):
        raise vol.Invalid("Nieprawidłowy adres e-mail.")
    return value


def _recipients(value):
    value = str(value).strip()
    items = split_recipients(value)
    if not items or not all(EMAIL_RE.match(i) for i in items):
        raise vol.Invalid("Nieprawidłowy adres odbiorcy.")
    return value


def _recipients_opt(value):
    if value is None or str(value).strip() == "":
        return ""
    return _recipients(value)


def _normalize_row(row: dict) -> dict:
    sheet = str(row.get(ROW_SHEET) or "").strip()
    if sheet:
        if row.get(ROW_ENTITY):
            raise vol.Invalid("Wiersz z arkuszem nie może mieć encji.")
        return {ROW_SHEET: sheet}
    out: dict = {}
    if row.get(ROW_ENTITY):
        out[ROW_ENTITY] = row[ROW_ENTITY]
    attr = str(row.get(ROW_ATTR) or "").strip()
    if attr:
        if not out:
            raise vol.Invalid("Atrybut wymaga encji.")
        out[ROW_ATTR] = attr
    desc = str(row.get(ROW_DESC) or "").strip()
    if desc:
        out[ROW_DESC] = desc
    if not out:
        raise vol.Invalid("Wiersz musi mieć encję, opis lub nazwę arkusza.")
    return out


def _unique(values):
    if len(set(values)) != len(values):
        raise vol.Invalid("Wartości nie mogą się powtarzać.")
    return values


def is_text_row(row: dict) -> bool:
    """Wiersz z samym opisem (bez encji i bez znacznika arkusza)."""
    return not row.get(ROW_ENTITY) and not row.get(ROW_SHEET)


def count_sheets(rows: list[dict]) -> int:
    markers = sum(1 for r in rows if r.get(ROW_SHEET))
    return markers + (0 if rows and rows[0].get(ROW_SHEET) else 1)


def _check_layout(conf: dict) -> dict:
    if conf.get(CONF_TEMPLATE):  # przy szablonie wiersze i kolumny nie są używane
        return conf
    if any(is_text_row(r) for r in conf[CONF_ROWS]) and "opis" not in conf[CONF_COLUMNS]:
        raise vol.Invalid(ERR_OPIS)
    return conf


def _rows_or_template(conf: dict) -> dict:
    if not conf.get(CONF_TEMPLATE) and not conf.get(CONF_ROWS):
        raise vol.Invalid("Raport wymaga wierszy (rows) albo szablonu (template).")
    return conf


def _template_path(value):
    value = str(value).strip()
    if not value.lower().endswith(".xlsx"):
        raise vol.Invalid("Szablon musi być plikiem .xlsx.")
    return value


def _set_report_id(report: dict) -> dict:
    report = dict(report)
    report.setdefault(CONF_ID, slugify(report[CONF_NAME]) or "raport")
    return report


def _unique_report_ids(reports: list[dict]) -> list[dict]:
    _unique([r[CONF_ID] for r in reports])
    return reports


def _legacy_yaml(conf):
    """Stary, płaski format YAML (jeden raport na poziomie bloku) -> lista raportów."""
    if isinstance(conf, dict) and CONF_REPORTS not in conf and CONF_ROWS in conf:
        conf = dict(conf)
        report = {
            k: conf.pop(k)
            for k in (CONF_CRON, CONF_RECIPIENT, CONF_SUBJECT, CONF_COLUMNS, CONF_ROWS)
            if k in conf
        }
        report.setdefault(CONF_NAME, "Raport")
        conf[CONF_REPORTS] = [report]
    return conf


ROW_SCHEMA = vol.All(
    vol.Schema(
        {
            vol.Optional(ROW_ENTITY): vol.Any(None, "", cv.entity_id),
            vol.Optional(ROW_ATTR): vol.Any(None, cv.string),
            vol.Optional(ROW_DESC): vol.Any(None, cv.string),
            vol.Optional(ROW_SHEET): vol.Any(None, cv.string),
        }
    ),
    _normalize_row,
)
COLUMNS_SCHEMA = vol.All(cv.ensure_list, [vol.In(list(COLUMN_LABELS))], vol.Length(min=1), _unique)
ROWS_SCHEMA = vol.All(cv.ensure_list, [ROW_SCHEMA], vol.Length(min=1))
LAYOUT_SCHEMA = vol.All(
    vol.Schema({vol.Required(CONF_COLUMNS): COLUMNS_SCHEMA, vol.Required(CONF_ROWS): ROWS_SCHEMA}),
    _check_layout,
)
# UWAGA: nowe klucze (cc, bcc, body) są opcjonalne i bez wartości domyślnych,
# żeby niezmieniony YAML dawał ten sam skrót i nie nadpisywał wpisu.
REPORT_SCHEMA = vol.All(
    vol.Schema(
        {
            vol.Optional(CONF_ID): cv.slug,
            vol.Required(CONF_NAME): cv.string,
            vol.Required(CONF_CRON): _cron,
            vol.Required(CONF_RECIPIENT): _recipients,
            vol.Optional(CONF_CC): _recipients_opt,
            vol.Optional(CONF_BCC): _recipients_opt,
            vol.Optional(CONF_SUBJECT, default=DEFAULT_SUBJECT): cv.string,
            vol.Optional(CONF_BODY): cv.string,
            vol.Optional(CONF_TEMPLATE): _template_path,
            vol.Optional(CONF_COLUMNS, default=DEFAULT_COLUMNS): COLUMNS_SCHEMA,
            vol.Optional(CONF_ROWS): ROWS_SCHEMA,
        }
    ),
    _rows_or_template,
    _check_layout,
    _set_report_id,
)
YAML_SCHEMA = vol.All(
    _legacy_yaml,
    vol.Schema(
        {
            vol.Optional(CONF_ID): cv.string,
            vol.Required(CONF_GMAIL_USER): _email,
            vol.Required(CONF_GMAIL_PASSWORD): cv.string,
            vol.Optional(CONF_RETRY, default=DEFAULT_RETRY): vol.All(
                vol.Coerce(int), vol.Range(min=1, max=1440)
            ),
            vol.Optional(CONF_RETENTION, default=DEFAULT_RETENTION): vol.All(
                vol.Coerce(int), vol.Range(min=0, max=3650)
            ),
            vol.Optional(CONF_STORAGE_PATH): cv.string,
            vol.Required(CONF_REPORTS): vol.All(
                cv.ensure_list, [REPORT_SCHEMA], vol.Length(min=1), _unique_report_ids
            ),
        }
    ),
)
CONFIG_SCHEMA = vol.Schema(
    {DOMAIN: vol.All(cv.ensure_list, [YAML_SCHEMA])}, extra=vol.ALLOW_EXTRA
)


def layout_error(columns, rows) -> str | None:
    """Zwraca opis błędu układu lub None, gdy układ jest poprawny."""
    try:
        LAYOUT_SCHEMA({CONF_COLUMNS: columns, CONF_ROWS: rows})
    except vol.Invalid as err:
        return str(err)
    return None


def yaml_digest(conf: dict) -> str:
    return hashlib.sha256(json.dumps(conf, sort_keys=True, default=str).encode()).hexdigest()


# ------------------------------------------------------------------- Excel ---
def _to_value(raw):
    try:
        number = float(raw)
    except (ValueError, TypeError):
        return raw
    return number if math.isfinite(number) else raw


def _attr_value(raw):
    """Wartość atrybutu -> typ możliwy do zapisania w komórce."""
    if raw is None:
        return ""
    if isinstance(raw, bool):
        return "tak" if raw else "nie"
    if isinstance(raw, (int, float)):
        return raw if math.isfinite(raw) else str(raw)
    if isinstance(raw, str):
        return _to_value(raw)
    return json.dumps(raw, ensure_ascii=False, default=str)


_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_TEXT_FILL = PatternFill("solid", fgColor="D9E1F2")
_BAND_FILL = PatternFill("solid", fgColor="F2F2F2")
_WARN_FILL = PatternFill("solid", fgColor="FFC7CE")
_ALIGN = {"lp": "center", "wartosc": "right", "jednostka": "center", "zmiana": "center"}
_DT_FORMAT = "yyyy-mm-dd hh:mm:ss"


def _sheet_title(title: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", " ", str(title)).strip()[:31] or "Arkusz"
    name, n = base, 2
    while name.lower() in used:
        suffix = f" {n}"
        name, n = base[: 31 - len(suffix)] + suffix, n + 1
    used.add(name.lower())
    return name


def _fill_sheet(ws, heading: str, columns, rows, snapshot: datetime) -> None:
    warns = sum(1 for r in rows if r["warn"])
    ws["A1"] = heading
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    note = f"Czas migawki: {snapshot.strftime(DT_FMT)}"
    if warns:
        note += f"   |   Encje niedostępne lub brakujące: {warns}"
    ws["A2"] = note
    ws["A2"].font = Font(italic=True, color="595959")

    widths = []
    for col, key in enumerate(columns, start=1):
        cell = ws.cell(row=FIRST_DATA_ROW - 1, column=col, value=COLUMN_LABELS[key])
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = _HEADER_FILL
        cell.border = _BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        widths.append(len(COLUMN_LABELS[key]))

    number = band = 0
    for i, row in enumerate(rows):
        r = FIRST_DATA_ROW + i
        values = dict(row["values"])
        is_text = row["kind"] == "text"
        if not is_text:
            number += 1
            band += 1
            values["lp"] = number
        for col, key in enumerate(columns, start=1):
            value = values.get(key)
            cell = ws.cell(row=r, column=col, value=value)
            if isinstance(value, str):
                cell.data_type = "s"  # tekst z encji nigdy nie jest traktowany jak formuła
            cell.border = _BORDER
            cell.alignment = Alignment(horizontal=_ALIGN.get(key, "left"), vertical="center")
            if key == "zmiana" and value is not None:
                cell.number_format = _DT_FORMAT
            if is_text:
                cell.fill = _TEXT_FILL
                cell.font = Font(bold=True)
                continue
            if band % 2 == 0:
                cell.fill = _BAND_FILL
            if row["warn"] and key == "wartosc":
                cell.fill = _WARN_FILL
                cell.font = Font(bold=True, color="9C0006")
            if value is not None:
                widths[col - 1] = max(widths[col - 1], len(str(value)))
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = min(max(width + 3, 8), 60)

    ws.freeze_panes = ws.cell(row=FIRST_DATA_ROW, column=1)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{FIRST_DATA_ROW - 1}:{FIRST_DATA_ROW - 1}"


def _build_xlsx(report_name: str, columns: list[str], rows: list[dict], snapshot: datetime) -> bytes:
    """Układ stały: wiersz N arkusza = N-ty wiersz raportu, kolumny wg `columns`.

    Wiersz z kluczem `sheet` rozpoczyna nowy arkusz (numeracja Lp. liczona od nowa).
    """
    sheets: list[tuple[str | None, list[dict]]] = []
    for row in rows:
        if row["kind"] == "sheet":
            sheets.append((row["values"]["name"], []))
        else:
            if not sheets:
                sheets.append((None, []))
            sheets[-1][1].append(row)
    if not sheets:
        sheets.append((None, []))

    wb = Workbook()
    wb.properties.title = report_name
    wb.properties.creator = "HA_mailer"
    used: set[str] = set()
    for idx, (title, sheet_rows) in enumerate(sheets):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = _sheet_title(title or DEFAULT_SHEET, used)
        heading = report_name if title is None else f"{report_name} - {title}"
        _fill_sheet(ws, heading, columns, sheet_rows, snapshot)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------ szablon Excel ---
def parse_token(token: str) -> dict | None:
    """Rozpoznaje znacznik z szablonu (zawartość między {{ i }}).

    Formy: `raport`, `data`, `czas`, `data_czas`, `ostrzezenia`, `domena.encja`,
    `domena.encja.atrybut`, z opcjonalnym przyrostkiem `#nazwa`, `#jednostka`, `#zmiana`.
    """
    token = token.strip()
    if token in SPECIAL_TOKENS:
        return {"kind": "special", "name": token}
    match = ENTITY_TOKEN_RE.match(token)
    if match:
        return {"kind": "entity", "entity": match[1], "attr": match[2], "mod": match[3]}
    return None


def _template_cells(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                # komórki z formułami pomijamy: znaczniki działają tylko w komórkach tekstowych
                if isinstance(value, str) and "{{" in value and not value.startswith("="):
                    yield cell, value


def scan_template(path: str):
    """Wczytuje szablon i zwraca (skoroszyt, unikalne znaczniki w kolejności występowania)."""
    wb = load_workbook(path)
    tokens: list[str] = []
    seen: set[str] = set()
    for _, value in _template_cells(wb):
        for token in TOKEN_RE.findall(value):
            if token not in seen:
                seen.add(token)
                tokens.append(token)
    return wb, tokens


def _as_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime(DT_FMT)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _set_cell(cell, value, warn: bool) -> None:
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"  # tekst z encji nigdy nie jest traktowany jak formuła
    if isinstance(value, datetime) and cell.number_format == "General":
        cell.number_format = _DT_FORMAT
    if warn:
        font = cell.font
        cell.fill = _WARN_FILL
        cell.font = Font(name=font.name, size=font.size, italic=font.italic, bold=True, color="9C0006")


def fill_template(wb, resolved: dict[str, dict]) -> bytes:
    """Podstawia wartości w miejsce znaczników, zachowując formatowanie szablonu.

    Komórka zawierająca wyłącznie znacznik dostaje wartość typowaną (liczba, data), a znacznik
    wewnątrz tekstu jest podmieniany na tekst. Nieznane znaczniki zostają bez zmian.
    """
    for cell, value in _template_cells(wb):
        whole = TOKEN_RE.fullmatch(value.strip())
        if whole and whole[1] in resolved:
            info = resolved[whole[1]]
            _set_cell(cell, info["value"], info["warn"])
        else:
            cell.value = TOKEN_RE.sub(
                lambda m: _as_text(resolved[m[1]]["value"]) if m[1] in resolved else m[0], value
            )
    wb.calculation.fullCalcOnLoad = True  # Excel przeliczy formuły szablonu po otwarciu pliku
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------- operacje blokujące (executor) ---
def _write_file(folder: str, name: str, payload: bytes) -> None:
    os.makedirs(folder, exist_ok=True)
    with open(os.path.join(folder, name), "wb") as fh:
        fh.write(payload)


def _read_file(path: str) -> bytes:
    with open(path, "rb") as fh:
        return fh.read()


def _cleanup_dirs(
    dirs: list[str], cutoff: float | None, protected: set[tuple[str, str]]
) -> tuple[int, int]:
    """Skanuje foldery i usuwa pliki raportów starsze niż `cutoff` (czas modyfikacji).

    Dotyka wyłącznie plików o nazwie utworzonej przez HA_mailer. Zwraca (usunięte, pozostałe).
    """
    deleted = remaining = 0
    for folder in dirs:
        try:
            names = os.listdir(folder)
        except OSError:
            continue
        for name in names:
            if not FILE_RE.match(name):
                continue
            if (folder, name) in protected:
                remaining += 1
                continue
            path = os.path.join(folder, name)
            try:
                st = os.stat(path)
                if not stat.S_ISREG(st.st_mode):
                    continue
                if cutoff is not None and st.st_mtime < cutoff:
                    os.remove(path)
                    deleted += 1
                else:
                    remaining += 1
            except OSError as err:
                _LOGGER.warning("HA_mailer: nie można usunąć %s: %s", name, err)
    return deleted, remaining


def _send_mail(cfg: dict, report: dict, filename: str, subject: str, payload: bytes, body: str) -> None:
    to = split_recipients(report[CONF_RECIPIENT])
    cc = split_recipients(report.get(CONF_CC))
    bcc = split_recipients(report.get(CONF_BCC))
    msg = EmailMessage()
    msg["From"] = cfg[CONF_GMAIL_USER]
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(payload, maintype=XLSX_MIME[0], subtype=XLSX_MIME[1], filename=filename)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=context, timeout=30) as smtp:
        smtp.login(cfg[CONF_GMAIL_USER], cfg[CONF_GMAIL_PASSWORD])
        smtp.send_message(msg, to_addrs=to + cc + bcc)  # UDW nie trafia do nagłówków


def _when(value: str | None, default: datetime) -> datetime:
    return (dt_util.parse_datetime(value) if value else None) or default


# ------------------------------------------------------------------ logika ---
class HAMailer:
    """Obsługa jednego wpisu: wiele raportów, kolejka wysyłki, archiwum plików."""

    def __init__(self, hass: HomeAssistant, entry: ConfigEntry) -> None:
        self.hass = hass
        self.entry = entry
        self.stats: dict[str, dict] = {}
        self.enabled: dict[str, bool] = {}  # przełącznik harmonogramu per raport
        self.pending: list[dict] = []  # pliki czekające na wysyłkę (świeże i zaległe)
        self.dirs: list[str] = []  # foldery, w których dodatek zapisywał pliki
        self.file_count = 0
        self.folder_ok = True
        self.next_runs: dict[str, datetime] = {}
        self._mem: dict[str, bytes] = {}  # pliki, których nie udało się zapisać na dysku
        self._store = Store(hass, 1, f"{DOMAIN}.{entry.entry_id}")
        self._unsubs: dict[str, Callable] = {}
        self._retry_unsub: Callable | None = None
        self._cleanup_unsub: Callable | None = None
        self._send_lock = asyncio.Lock()
        self._flush_task: asyncio.Task | None = None
        self._tasks: set[asyncio.Task] = set()
        self._stopped = False

    # ---- dostęp do konfiguracji
    @property
    def cfg(self) -> dict:
        return {**self.entry.data, **self.entry.options}

    @property
    def reports(self) -> dict[str, dict]:
        return {r[CONF_ID]: r for r in self.cfg[CONF_REPORTS]}

    @property
    def folder(self) -> str:
        return self.hass.config.path(self.cfg.get(CONF_STORAGE_PATH) or DEFAULT_FOLDER)

    @property
    def retry_interval(self) -> timedelta:
        return timedelta(minutes=int(self.cfg.get(CONF_RETRY, DEFAULT_RETRY)))

    def find_report(self, text: str) -> str | None:
        needle = str(text).strip().lower()
        for rid, report in self.reports.items():
            if needle in (rid.lower(), str(report[CONF_NAME]).strip().lower()):
                return rid
        return None

    def pending_for(self, rid: str) -> list[dict]:
        return [i for i in self.pending if i["report"] == rid]

    def device_info(self, rid: str | None = None) -> DeviceInfo:
        eid = self.entry.entry_id
        if rid is None:
            return DeviceInfo(identifiers={(DOMAIN, eid)}, name="HA_mailer", manufacturer="HA_mailer")
        return DeviceInfo(
            identifiers={(DOMAIN, f"{eid}_{rid}")},
            name=self.reports[rid][CONF_NAME],
            manufacturer="HA_mailer",
            via_device=(DOMAIN, eid),
        )

    def expected_unique_ids(self) -> set[str]:
        eid = self.entry.entry_id
        ids = {f"{eid}_pending"}
        for rid in self.reports:
            ids |= {f"{eid}_{rid}_{k}" for k in ("status", "last_sent", "stats", "enabled")}
        return ids

    # ---- stan trwały
    async def async_load(self) -> None:
        saved = await self._store.async_load() or {}
        ids = set(self.reports)
        for rid in ids:
            self.stats[rid] = {**DEFAULT_STATS, **saved.get("stats", {}).get(rid, {})}
            self.enabled[rid] = bool(saved.get("enabled", {}).get(rid, True))
        for gone in set(saved.get("stats", {})) - ids:  # raporty usunięte z konfiguracji
            self._clear_alerts(gone)
        self.pending = [i for i in saved.get("pending", []) if i["report"] in ids]
        self.dirs = list(saved.get("dirs", []))

    async def _save(self) -> None:
        snapshot = json.loads(
            json.dumps(
                {
                    "stats": self.stats,
                    "enabled": self.enabled,
                    "pending": self.pending,
                    "dirs": self.dirs,
                },
                default=str,
            )
        )
        await self._store.async_save(snapshot)  # kopia: brak wyścigu z modyfikacją stanu

    def _notify(self) -> None:
        async_dispatcher_send(self.hass, SIGNAL.format(self.entry.entry_id))

    # ---- cykl życia i harmonogramy
    def start(self) -> None:
        self._stopped = False
        for rid in self.reports:
            self._schedule_report(rid)
        self._schedule_retry()
        self._cleanup_unsub = async_track_time_interval(
            self.hass, self._cleanup, timedelta(hours=1)
        )
        self.hass.async_create_task(self._cleanup())

    def stop(self) -> None:
        """Synchronicznie: zatrzymaj timery i nie zakładaj nowych (idempotentne)."""
        self._stopped = True
        for rid in list(self._unsubs):
            self._cancel_report(rid)
        if self._retry_unsub:
            self._retry_unsub()
            self._retry_unsub = None
        if self._cleanup_unsub:
            self._cleanup_unsub()
            self._cleanup_unsub = None

    async def async_shutdown(self) -> None:
        """Zatrzymaj i poczekaj na zakończenie trwających operacji (bez duplikatów po reloadzie)."""
        self.stop()
        tasks = [t for t in self._tasks if not t.done()]
        if tasks:
            _, still_running = await asyncio.wait(tasks, timeout=90)
            for task in still_running:
                task.cancel()

    def _cancel_report(self, rid: str) -> None:
        unsub = self._unsubs.pop(rid, None)
        if unsub:
            unsub()

    def _schedule_report(self, rid: str) -> None:
        self._cancel_report(rid)
        report = self.reports.get(rid)
        if self._stopped or not report or not self.enabled.get(rid, True):
            self.next_runs.pop(rid, None)
            self._notify()
            return
        nxt = croniter(report[CONF_CRON], dt_util.now()).get_next(datetime)
        self.next_runs[rid] = nxt
        self._unsubs[rid] = async_track_point_in_time(
            self.hass, partial(self._fire_report, rid), nxt
        )
        self._notify()

    async def async_set_enabled(self, rid: str, value: bool) -> None:
        self.enabled[rid] = bool(value)
        await self._save()
        self._schedule_report(rid)

    async def _fire_report(self, rid: str, _now: datetime) -> None:
        """Termin z harmonogramu: utwórz plik (w tle) i uruchom wysyłkę zaległych."""
        me = asyncio.current_task()
        self._tasks.add(me)
        self._unsubs.pop(rid, None)
        try:
            if self.enabled.get(rid, True):
                await self._create_scheduled(rid)
                self._kick_flush()
        except HomeAssistantError:
            pass  # błąd zarejestrowany (log, powiadomienie, encje)
        except Exception:  # noqa: BLE001
            _LOGGER.exception("Nieoczekiwany błąd harmonogramu HA_mailer (%s)", rid)
        finally:
            self._tasks.discard(me)
            self._schedule_report(rid)

    def _schedule_retry(self) -> None:
        if self._retry_unsub:
            self._retry_unsub()
            self._retry_unsub = None
        if not self._stopped:
            times = [_when(i.get("next_try"), dt_util.now()) for i in self.pending]
            if times:
                self._retry_unsub = async_track_point_in_time(self.hass, self._retry_due, min(times))
        self._notify()

    async def _retry_due(self, _now: datetime) -> None:
        self._retry_unsub = None
        self._kick_flush()

    # ---- zbieranie danych do raportu
    def _collect_rows(self, report: dict) -> list[dict]:
        """Kolejność = kolejność wierszy w konfiguracji raportu (nigdy się nie zmienia)."""
        out: list[dict] = []
        for row in report.get(CONF_ROWS, []):
            if row.get(ROW_SHEET):
                out.append({"kind": "sheet", "values": {"name": row[ROW_SHEET]}, "warn": False})
                continue
            desc = row.get(ROW_DESC) or None
            entity_id = row.get(ROW_ENTITY)
            if not entity_id:
                out.append({"kind": "text", "values": {"opis": desc}, "warn": False})
                continue
            attr = row.get(ROW_ATTR)
            state = self.hass.states.get(entity_id)
            values = {"opis": desc, "encja": entity_id, "atrybut": attr}
            if state is None:
                values.update(nazwa=entity_id, wartosc="brak encji", jednostka="")
                warn = True
            else:
                values["nazwa"] = f"{state.name} - {attr}" if attr else state.name
                values["zmiana"] = dt_util.as_local(state.last_changed).replace(
                    tzinfo=None, microsecond=0
                )
                warn = state.state in WARN_STATES
                if attr:
                    if attr in state.attributes:
                        values.update(wartosc=_attr_value(state.attributes[attr]), jednostka="")
                    else:
                        values.update(wartosc="brak atrybutu", jednostka="")
                        warn = True
                else:
                    values.update(
                        wartosc=_to_value(state.state),
                        jednostka=state.attributes.get("unit_of_measurement", ""),
                    )
            out.append({"kind": "entity", "values": values, "warn": warn})
        return out

    # ---- tworzenie plików
    def _resolve_entity(self, info: dict) -> dict:
        """Wartość znacznika encji: {"value": ..., "warn": ...}."""
        eid, attr, mod = info["entity"], info["attr"], info["mod"]
        state = self.hass.states.get(eid)
        if state is None:
            value = {"nazwa": eid, "jednostka": "", "zmiana": ""}.get(mod, "brak encji")
            return {"value": value, "warn": True}
        if mod == "nazwa":
            return {"value": state.name, "warn": False}
        if mod == "jednostka":
            return {"value": state.attributes.get("unit_of_measurement", ""), "warn": False}
        if mod == "zmiana":
            changed = dt_util.as_local(state.last_changed).replace(tzinfo=None, microsecond=0)
            return {"value": changed, "warn": False}
        if attr:
            if attr in state.attributes:
                return {"value": _attr_value(state.attributes[attr]), "warn": state.state in WARN_STATES}
            return {"value": "brak atrybutu", "warn": True}
        return {"value": _to_value(state.state), "warn": state.state in WARN_STATES}

    async def _build_from_template(self, report: dict, now: datetime) -> tuple[bytes, int, int]:
        """Wypełnij szablon .xlsx wartościami encji. Zwraca (plik, liczba_encji, ostrzeżenia)."""
        path = self.hass.config.path(report[CONF_TEMPLATE])
        wb, tokens = await self.hass.async_add_executor_job(scan_template, path)
        resolved: dict[str, dict] = {}
        entity_ids: set[str] = set()
        problems: set[tuple[str, str | None]] = set()
        for token in tokens:
            info = parse_token(token)
            if info and info["kind"] == "entity":
                resolved[token] = self._resolve_entity(info)
                entity_ids.add(info["entity"])
                if resolved[token]["warn"]:
                    problems.add((info["entity"], info["attr"]))
        stamp = now.replace(tzinfo=None, microsecond=0)
        specials = {
            "raport": report[CONF_NAME],
            "data": stamp.strftime("%Y-%m-%d"),
            "czas": stamp.strftime("%H:%M:%S"),
            "data_czas": stamp,
            "ostrzezenia": len(problems),
        }
        for token in tokens:
            info = parse_token(token)
            if info and info["kind"] == "special":
                resolved[token] = {"value": specials[info["name"]], "warn": False}
            elif info is None:
                _LOGGER.warning("HA_mailer [%s]: nieznany znacznik {{%s}} w szablonie", report[CONF_NAME], token)
        payload = await self.hass.async_add_executor_job(fill_template, wb, resolved)
        return payload, len(entity_ids), len(problems)

    async def _create_item(self, rid: str) -> tuple[dict, bytes]:
        """Zbuduj plik z aktualnych stanów i zapisz go w folderze (lub w pamięci przy awarii dysku)."""
        report = self.reports[rid]
        now = dt_util.now()
        folder = self.folder
        try:
            if report.get(CONF_TEMPLATE):
                payload, entities, warnings = await self._build_from_template(report, now)
            else:
                rows = self._collect_rows(report)
                payload = await self.hass.async_add_executor_job(
                    _build_xlsx,
                    report[CONF_NAME],
                    list(report[CONF_COLUMNS]),
                    rows,
                    now.replace(tzinfo=None, microsecond=0),
                )
                entities = sum(1 for r in rows if r["kind"] == "entity")
                warnings = sum(1 for r in rows if r["warn"])
        except Exception as err:  # noqa: BLE001
            error = f"Nie można utworzyć pliku raportu: {err}"
            await self._record(rid, now, error)
            raise HomeAssistantError(f"HA_mailer: {error}") from err
        name = f"HA_mailer_{rid}_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        item = {
            "id": uuid.uuid4().hex,
            "report": rid,
            "dir": folder,
            "file": name,
            "stored": False,
            "subject": f"{report[CONF_SUBJECT]} {now.strftime(DT_FMT)}",
            "created": now.isoformat(),
            "entities": entities,
            "warnings": warnings,
            "attempts": 0,
            "last_error": None,
            "next_try": None,
        }
        st = self.stats.setdefault(rid, dict(DEFAULT_STATS))
        st["last_entities"], st["last_warnings"] = entities, warnings
        try:
            await self.hass.async_add_executor_job(_write_file, folder, name, payload)
        except OSError as err:
            self._mem[item["id"]] = payload  # raport nie przepadnie: wyślemy go z pamięci
            self._storage_problem(folder, err)
        else:
            item["stored"] = True
            if folder not in self.dirs:
                self.dirs.append(folder)
            self._storage_ok()
        return item, payload

    async def _create_scheduled(self, rid: str) -> None:
        item, _ = await self._create_item(rid)
        item["next_try"] = item["created"]  # do wysłania od razu
        self.pending.append(item)
        await self._save()
        self._notify()

    # ---- alerty o folderze
    def _storage_problem(self, folder: str, err: Exception) -> None:
        self.folder_ok = False
        _LOGGER.warning("HA_mailer: nie można zapisać pliku w %s: %s", folder, err)
        ir.async_create_issue(
            self.hass,
            DOMAIN,
            f"storage_unavailable_{self.entry.entry_id}",
            is_fixable=False,
            severity=ir.IssueSeverity.WARNING,
            translation_key="storage_unavailable",
            translation_placeholders={"path": folder, "error": str(err)},
        )

    def _storage_ok(self) -> None:
        if not self.folder_ok:
            self.folder_ok = True
            ir.async_delete_issue(self.hass, DOMAIN, f"storage_unavailable_{self.entry.entry_id}")

    def _clear_alerts(self, rid: str) -> None:
        persistent_notification.async_dismiss(self.hass, f"{DOMAIN}_{self.entry.entry_id}_{rid}")
        ir.async_delete_issue(self.hass, DOMAIN, f"send_failed_{self.entry.entry_id}_{rid}")

    # ---- rejestr wyników
    async def _record(
        self, rid: str, now: datetime, error: str | None, size_kb: float | None = None, retried: bool = False
    ) -> None:
        st = self.stats.setdefault(rid, dict(DEFAULT_STATS))
        name = self.reports.get(rid, {}).get(CONF_NAME, rid)
        st["last_attempt"] = now.isoformat()
        if error is None:
            st["ok"] += 1
            st["status"] = "ok"
            st["last_ok"] = now.isoformat()
            st["last_file_kb"] = size_kb
            if retried:
                st["retry_ok"] += 1
            if not self.pending_for(rid):
                self._clear_alerts(rid)
        else:
            st["fail"] += 1
            st["status"] = "error"
            st["last_error"] = error
            st["last_error_time"] = now.isoformat()
            _LOGGER.error("HA_mailer [%s]: wysyłka nieudana: %s", name, error)
            retry_note = ""
            if self.pending_for(rid):
                minutes = int(self.retry_interval.total_seconds() // 60)
                retry_note = (
                    f"\n\nPliki czekają w kolejce, a wysyłka będzie ponawiana co {minutes} min "
                    f"(w kolejce: {len(self.pending_for(rid))})."
                )
            persistent_notification.async_create(
                self.hass,
                f"Raport „{name}” - błąd wysyłki ({now.strftime(DT_FMT)}):\n\n{error}{retry_note}",
                title="HA_mailer - błąd wysyłki",
                notification_id=f"{DOMAIN}_{self.entry.entry_id}_{rid}",
            )
            ir.async_create_issue(
                self.hass,
                DOMAIN,
                f"send_failed_{self.entry.entry_id}_{rid}",
                is_fixable=False,
                severity=ir.IssueSeverity.ERROR,
                translation_key="send_failed",
                translation_placeholders={"report": name, "error": error},
            )
        await self._save()
        self._notify()

    # ---- wysyłka
    async def _deliver(self, item: dict, payload: bytes) -> tuple[str | None, str | None]:
        """Wysyła plik z `item`. Zwraca (opis_błędu, zakres) - zakres: "item" lub "global"."""
        report = self.reports[item["report"]]
        created = _when(item["created"], dt_util.now())
        created_txt = dt_util.as_local(created).strftime(DT_FMT)
        warnings = int(item.get("warnings", 0))
        custom = (report.get(CONF_BODY) or "").strip()
        if custom:
            body = (
                custom.replace("[raport]", str(report[CONF_NAME]))
                .replace("[data]", created_txt)
                .replace("[liczba_encji]", str(item["entities"]))
                .replace("[ostrzezenia]", str(warnings))
            )
        else:
            body = f"Raport „{report[CONF_NAME]}” z {created_txt}. W załączniku wartości {item['entities']} encji."
            if warnings:
                body += f"\nUwaga: encje niedostępne lub brakujące: {warnings} (zaznaczone w pliku)."
        if item["attempts"] > 0:
            body += (
                f"\n\nWysyłka była opóźniona (nieudane próby: {item['attempts']}), "
                f"ponowiona {dt_util.now().strftime(DT_FMT)}. Plik zawiera stan z chwili utworzenia raportu."
            )
        try:
            await self.hass.async_add_executor_job(
                _send_mail, self.cfg, report, item["file"], item["subject"], payload, body
            )
        except smtplib.SMTPAuthenticationError:
            return "Gmail odrzucił logowanie - sprawdź adres konta i hasło aplikacji.", "global"
        except smtplib.SMTPRecipientsRefused:
            return "Serwer odrzucił adres odbiorcy - sprawdź adresy raportu.", "item"
        except (OSError, smtplib.SMTPException) as err:
            return f"Błąd połączenia/SMTP: {err}", "global"
        except Exception as err:  # noqa: BLE001
            return f"Nieoczekiwany błąd: {err!r}", "global"
        return None, None

    def _kick_flush(self) -> None:
        """Uruchom w tle wysyłanie wszystkich zaległych plików (jedna pętla naraz)."""
        if self._stopped or (self._flush_task and not self._flush_task.done()):
            return
        self._flush_task = self.hass.async_create_task(self._flush())

    async def _flush(self) -> None:
        me = asyncio.current_task()
        self._tasks.add(me)
        try:
            async with self._send_lock:
                while not self._stopped:
                    limit = dt_util.now() + timedelta(seconds=1)
                    due = sorted(
                        (i for i in self.pending if _when(i.get("next_try"), limit) <= limit),
                        key=lambda i: i["created"],
                    )
                    if not due:
                        break
                    if await self._send_item(due[0]) == "global":
                        # sieć lub logowanie: nie męcz reszty, spróbuj za interwał
                        later = (dt_util.now() + self.retry_interval).isoformat()
                        for other in due[1:]:
                            other["next_try"] = later
                        await self._save()
                        break
        except Exception:  # noqa: BLE001
            _LOGGER.exception("Nieoczekiwany błąd wysyłania kolejki HA_mailer")
        finally:
            self._tasks.discard(me)
            self._schedule_retry()

    async def _load_payload(self, item: dict) -> bytes | None:
        if item["id"] in self._mem:
            return self._mem[item["id"]]
        try:
            return await self.hass.async_add_executor_job(
                _read_file, os.path.join(item["dir"], item["file"])
            )
        except OSError:
            return None

    async def _send_item(self, item: dict) -> str | None:
        """Jedna próba wysyłki pliku z kolejki. Zwraca zakres błędu lub None."""
        now = dt_util.now()
        rid = item["report"]
        if rid not in self.reports:  # raport usunięty z konfiguracji
            self.pending.remove(item)
            self._mem.pop(item["id"], None)
            await self._save()
            return None
        payload = await self._load_payload(item)
        if payload is None:
            self.pending.remove(item)
            await self._record(
                rid, now, f"Plik {item['file']} jest niedostępny - porzucono jego wysyłkę."
            )
            return "item"
        if not item.get("stored", True):  # spróbuj ponownie zapisać plik na dysku
            try:
                await self.hass.async_add_executor_job(
                    _write_file, item["dir"], item["file"], payload
                )
            except OSError:
                pass
            else:
                item["stored"] = True
                self._mem.pop(item["id"], None)
                if item["dir"] not in self.dirs:
                    self.dirs.append(item["dir"])
                self._storage_ok()
        try:
            error, scope = await self._deliver(item, payload)
        except Exception as err:  # noqa: BLE001
            error, scope = f"Nieoczekiwany błąd: {err!r}", "item"
        if error is None:
            retried = item["attempts"] > 0
            self.pending.remove(item)
            self._mem.pop(item["id"], None)
            await self._record(rid, now, None, round(len(payload) / 1024, 1), retried=retried)
            return None
        item["attempts"] += 1
        item["last_error"] = error
        item["next_try"] = (now + self.retry_interval).isoformat()
        await self._record(rid, now, error)
        return scope

    async def async_send_now(self, rid: str) -> None:
        """Ręczna wysyłka: bez kolejki, błąd wraca do wywołującego."""
        item, payload = await self._create_item(rid)
        async with self._send_lock:
            now = dt_util.now()
            error, _ = await self._deliver(item, payload)
        self._mem.pop(item["id"], None)
        if error is None:
            await self._record(rid, now, None, round(len(payload) / 1024, 1))
            return
        await self._record(rid, now, error)
        raise HomeAssistantError(f"HA_mailer: {error}")

    async def async_clear_pending(self, rids: list[str] | None = None) -> int:
        """Porzuć niewysłane pliki z kolejki (zostają w folderze do końca retencji)."""
        async with self._send_lock:
            before = len(self.pending)
            for item in [i for i in self.pending if rids is None or i["report"] in rids]:
                self.pending.remove(item)
                self._mem.pop(item["id"], None)
            for rid in rids if rids is not None else list(self.reports):
                if not self.pending_for(rid):
                    self._clear_alerts(rid)
            await self._save()
            self._schedule_retry()
            return before - len(self.pending)

    # ---- archiwum plików
    async def _cleanup(self, _now: datetime | None = None) -> None:
        """Skanuje foldery i usuwa pliki starsze niż retencja (wg daty zapisu pliku)."""
        me = asyncio.current_task()
        self._tasks.add(me)
        try:
            days = int(self.cfg.get(CONF_RETENTION, DEFAULT_RETENTION))
            cutoff = (dt_util.now() - timedelta(days=days)).timestamp() if days > 0 else None
            dirs = list(dict.fromkeys([*self.dirs, self.folder]))
            protected = {(i["dir"], i["file"]) for i in self.pending}
            deleted, remaining = await self.hass.async_add_executor_job(
                _cleanup_dirs, dirs, cutoff, protected
            )
            self.file_count = remaining
            if deleted:
                _LOGGER.debug("HA_mailer: usunięto %s starych plików", deleted)
            self._notify()
        except Exception:  # noqa: BLE001
            _LOGGER.exception("Błąd czyszczenia archiwum HA_mailer")
        finally:
            self._tasks.discard(me)


# ------------------------------------------------------------- cykl życia ---
async def async_setup(hass: HomeAssistant, config: dict) -> bool:
    """Import konfiguracji z configuration.yaml do wpisów w UI."""
    for idx, conf in enumerate(config.get(DOMAIN, [])):
        conf = dict(conf)
        conf.setdefault(CONF_ID, str(idx))
        hass.async_create_task(
            hass.config_entries.flow.async_init(
                DOMAIN, context={"source": SOURCE_IMPORT}, data=conf
            )
        )
    return True


async def async_migrate_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """v1 (lista encji) i v2 (jeden raport) -> v3 (wiele raportów, wspólne ustawienia)."""
    if entry.version >= 3:
        return True

    def to_rows(src) -> dict:
        d = dict(src)
        if CONF_ENTITIES in d:
            d[CONF_ROWS] = [{ROW_ENTITY: e} for e in d.pop(CONF_ENTITIES)]
        return d

    data, options = to_rows(entry.data), to_rows(entry.options)
    if entry.version == 1:
        data.setdefault(CONF_COLUMNS, list(LEGACY_COLUMNS))

    def build(src: dict) -> dict:
        report = {
            CONF_ID: "raport",
            CONF_NAME: "Raport",
            CONF_CRON: src[CONF_CRON],
            CONF_RECIPIENT: src[CONF_RECIPIENT],
            CONF_SUBJECT: src.get(CONF_SUBJECT, DEFAULT_SUBJECT),
            CONF_COLUMNS: list(src.get(CONF_COLUMNS, LEGACY_COLUMNS)),
            CONF_ROWS: list(src[CONF_ROWS]),
        }
        return {
            CONF_GMAIL_USER: src[CONF_GMAIL_USER],
            CONF_GMAIL_PASSWORD: src[CONF_GMAIL_PASSWORD],
            CONF_RETRY: DEFAULT_RETRY,
            CONF_RETENTION: DEFAULT_RETENTION,
            CONF_REPORTS: [report],
        }

    new = build({**data, **options})
    if CONF_YAML_HASH in data:
        # wpis z YAML: zachowaj ID i skrót tak, aby niezmieniony YAML nie nadpisał wpisu
        new[CONF_ID] = data.get(CONF_ID, "0")
        new[CONF_YAML_HASH] = yaml_digest({**build(data), CONF_ID: new[CONF_ID]})

    store = Store(hass, 1, f"{DOMAIN}.{entry.entry_id}")
    old = await store.async_load()
    if old and "stats" not in old:  # stare statystyki -> statystyki raportu "raport"
        await store.async_save({"stats": {"raport": old}, "pending": []})

    @callback
    def _migrate_unique_id(ent):
        suffix = ent.unique_id.removeprefix(f"{entry.entry_id}_")
        if suffix in ("status", "last_sent", "stats"):
            return {"new_unique_id": f"{entry.entry_id}_raport_{suffix}"}
        return None

    await er.async_migrate_entries(hass, entry.entry_id, _migrate_unique_id)
    hass.config_entries.async_update_entry(entry, data=new, options={}, version=3)
    return True


def _prune_registry(hass: HomeAssistant, entry: ConfigEntry, mailer: HAMailer) -> None:
    """Usuń encje i urządzenia raportów, których już nie ma w konfiguracji."""
    expected = mailer.expected_unique_ids()
    ent_reg = er.async_get(hass)
    for ent in er.async_entries_for_config_entry(ent_reg, entry.entry_id):
        if ent.unique_id not in expected:
            ent_reg.async_remove(ent.entity_id)
    expected_dev = {(DOMAIN, entry.entry_id)} | {
        (DOMAIN, f"{entry.entry_id}_{rid}") for rid in mailer.reports
    }
    dev_reg = dr.async_get(hass)
    for dev in dr.async_entries_for_config_entry(dev_reg, entry.entry_id):
        if not (dev.identifiers & expected_dev):
            dev_reg.async_remove_device(dev.id)


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    mailer = HAMailer(hass, entry)
    await mailer.async_load()
    hass.data.setdefault(DOMAIN, {})[entry.entry_id] = mailer
    _prune_registry(hass, entry, mailer)
    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)
    mailer.start()
    entry.async_on_unload(mailer.stop)
    entry.async_on_unload(entry.add_update_listener(_async_reload))
    _register_services(hass)
    return True


def _targets(hass: HomeAssistant, call: ServiceCall) -> list[tuple[HAMailer, list[str]]]:
    mailers: dict[str, HAMailer] = hass.data.get(DOMAIN, {})
    entry_id = call.data.get("entry_id")
    if entry_id and entry_id not in mailers:
        raise ServiceValidationError(f"Nieznany entry_id: {entry_id}")
    selected = [mailers[entry_id]] if entry_id else list(mailers.values())
    text = call.data.get("report")
    result: list[tuple[HAMailer, list[str]]] = []
    for mailer in selected:
        if text:
            rid = mailer.find_report(text)
            if rid:
                result.append((mailer, [rid]))
        else:
            result.append((mailer, list(mailer.reports)))
    if text and not result:
        raise ServiceValidationError(f"Nie znaleziono raportu: {text}")
    return result


def _register_services(hass: HomeAssistant) -> None:
    if hass.services.has_service(DOMAIN, "send_now"):
        return
    schema = vol.Schema({vol.Optional("entry_id"): str, vol.Optional("report"): str})

    async def _send_now(call: ServiceCall) -> None:
        errors: list[str] = []
        for mailer, rids in _targets(hass, call):
            for rid in rids:
                try:
                    await mailer.async_send_now(rid)
                except HomeAssistantError as err:
                    errors.append(str(err))
        if errors:
            raise HomeAssistantError("; ".join(errors))

    async def _clear_pending(call: ServiceCall) -> None:
        for mailer, rids in _targets(hass, call):
            await mailer.async_clear_pending(rids)

    hass.services.async_register(DOMAIN, "send_now", _send_now, schema=schema)
    hass.services.async_register(DOMAIN, "clear_pending", _clear_pending, schema=schema)


async def _async_reload(hass: HomeAssistant, entry: ConfigEntry) -> None:
    await hass.config_entries.async_reload(entry.entry_id)


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    mailer: HAMailer | None = hass.data.get(DOMAIN, {}).get(entry.entry_id)
    if mailer:
        await mailer.async_shutdown()
    ok = await hass.config_entries.async_unload_platforms(entry, PLATFORMS)
    if ok:
        hass.data[DOMAIN].pop(entry.entry_id, None)
        if not hass.data[DOMAIN]:
            hass.services.async_remove(DOMAIN, "send_now")
            hass.services.async_remove(DOMAIN, "clear_pending")
    return ok


async def async_remove_entry(hass: HomeAssistant, entry: ConfigEntry) -> None:
    store = Store(hass, 1, f"{DOMAIN}.{entry.entry_id}")
    saved = await store.async_load() or {}
    for rid in saved.get("stats", {}):
        persistent_notification.async_dismiss(hass, f"{DOMAIN}_{entry.entry_id}_{rid}")
        ir.async_delete_issue(hass, DOMAIN, f"send_failed_{entry.entry_id}_{rid}")
    ir.async_delete_issue(hass, DOMAIN, f"storage_unavailable_{entry.entry_id}")
    await store.async_remove()
