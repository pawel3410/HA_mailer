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
import uuid
from collections.abc import Callable
from datetime import datetime, timedelta
from email.message import EmailMessage
from functools import partial
from io import BytesIO

import voluptuous as vol
from croniter import croniter
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from homeassistant.components import persistent_notification
from homeassistant.config_entries import SOURCE_IMPORT, ConfigEntry
from homeassistant.const import CONF_NAME
from homeassistant.core import HomeAssistant, ServiceCall, callback
from homeassistant.exceptions import HomeAssistantError, ServiceValidationError
import homeassistant.helpers.config_validation as cv
from homeassistant.helpers import entity_registry as er
from homeassistant.helpers.dispatcher import async_dispatcher_send
from homeassistant.helpers.event import async_track_point_in_time, async_track_time_interval
from homeassistant.helpers.storage import Store
from homeassistant.util import dt as dt_util, slugify

_LOGGER = logging.getLogger(__name__)

DOMAIN = "ha_mailer"
PLATFORMS = ["sensor"]

CONF_ID = "id"
CONF_ENTITIES = "entities"  # tylko w kroku początkowym UI (zamieniane na wiersze)
CONF_CRON = "cron"
CONF_GMAIL_USER = "gmail_user"
CONF_GMAIL_PASSWORD = "gmail_app_password"
CONF_RECIPIENT = "recipient"
CONF_SUBJECT = "subject"
CONF_COLUMNS = "columns"
CONF_ROWS = "rows"
CONF_REPORTS = "reports"
CONF_RETRY = "retry_minutes"
CONF_RETENTION = "retention_days"
CONF_STORAGE_PATH = "storage_path"
CONF_YAML_HASH = "yaml_hash"

ROW_ENTITY = "entity"
ROW_DESC = "description"

# klucz kolumny -> nagłówek w Excelu (kolejność = domyślna kolejność)
COLUMN_LABELS = {
    "lp": "Lp.",
    "opis": "Opis",
    "encja": "Encja",
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
GMAIL_HOST = "smtp.gmail.com"
GMAIL_PORT = 465

SIGNAL = DOMAIN + "_update_{}"
DT_FMT = "%Y-%m-%d %H:%M:%S"
XLSX_MIME = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
FIRST_DATA_ROW = 4  # wiersz 1: czas migawki, wiersz 3: nagłówki, dane od wiersza 4

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
}


# ---------------------------------------------------------------- walidacja ---
def split_recipients(text: str) -> list[str]:
    return [r.strip() for r in str(text).replace(";", ",").split(",") if r.strip()]


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


def _normalize_row(row: dict) -> dict:
    out: dict = {}
    if row.get(ROW_ENTITY):
        out[ROW_ENTITY] = row[ROW_ENTITY]
    desc = str(row.get(ROW_DESC) or "").strip()
    if desc:
        out[ROW_DESC] = desc
    if not out:
        raise vol.Invalid("Wiersz musi mieć encję lub opis.")
    return out


def _unique(values):
    if len(set(values)) != len(values):
        raise vol.Invalid("Wartości nie mogą się powtarzać.")
    return values


def _check_layout(conf: dict) -> dict:
    if any(not r.get(ROW_ENTITY) for r in conf[CONF_ROWS]) and "opis" not in conf[CONF_COLUMNS]:
        raise vol.Invalid(ERR_OPIS)
    return conf


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
            vol.Optional(ROW_DESC): vol.Any(None, cv.string),
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
REPORT_SCHEMA = vol.All(
    vol.Schema(
        {
            vol.Optional(CONF_ID): cv.slug,
            vol.Required(CONF_NAME): cv.string,
            vol.Required(CONF_CRON): _cron,
            vol.Required(CONF_RECIPIENT): _recipients,
            vol.Optional(CONF_SUBJECT, default=DEFAULT_SUBJECT): cv.string,
            vol.Optional(CONF_COLUMNS, default=DEFAULT_COLUMNS): COLUMNS_SCHEMA,
            vol.Required(CONF_ROWS): ROWS_SCHEMA,
        }
    ),
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
def _to_value(raw: str):
    try:
        number = float(raw)
    except (ValueError, TypeError):
        return raw
    return number if math.isfinite(number) else raw


def _build_xlsx(columns: list[str], rows: list[tuple[bool, dict]], snapshot: datetime) -> bytes:
    """Układ stały: wiersz N pliku = N-ty wiersz z konfiguracji, kolumny wg `columns`."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Migawka"
    ws["A1"] = "Czas migawki:"
    ws["A1"].font = Font(bold=True)
    ws["B1"] = snapshot
    ws["B1"].number_format = "yyyy-mm-dd hh:mm:ss"
    ws["B1"].alignment = Alignment(horizontal="left")
    widths = []
    for col, key in enumerate(columns, start=1):
        cell = ws.cell(row=3, column=col, value=COLUMN_LABELS[key])
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
        widths.append(len(COLUMN_LABELS[key]))
    text_fill = PatternFill("solid", fgColor="D9E1F2")
    for i, (is_text, values) in enumerate(rows):
        r = FIRST_DATA_ROW + i
        for col, key in enumerate(columns, start=1):
            value = values.get(key)
            cell = ws.cell(row=r, column=col, value=value)
            if isinstance(value, str):
                cell.data_type = "s"  # tekst z encji nigdy nie jest traktowany jak formuła
            if key == "zmiana" and value is not None:
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            if is_text:
                cell.fill = text_fill
                cell.font = Font(bold=True)
            elif value is not None:
                widths[col - 1] = max(widths[col - 1], len(str(value)))
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = min(width + 3, 60)
    ws.freeze_panes = "A4"
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


def _delete_files(items: list[tuple[str, str]]) -> None:
    for folder, name in items:
        if os.path.basename(name) != name:  # tylko pliki utworzone przez dodatek
            continue
        try:
            os.remove(os.path.join(folder, name))
        except FileNotFoundError:
            pass
        except OSError as err:
            _LOGGER.warning("HA_mailer: nie można usunąć %s: %s", name, err)


def _send_mail(cfg: dict, recipient: str, filename: str, subject: str, payload: bytes, body: str) -> None:
    msg = EmailMessage()
    msg["From"] = cfg[CONF_GMAIL_USER]
    msg["To"] = ", ".join(split_recipients(recipient))
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(payload, maintype=XLSX_MIME[0], subtype=XLSX_MIME[1], filename=filename)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=context, timeout=30) as smtp:
        smtp.login(cfg[CONF_GMAIL_USER], cfg[CONF_GMAIL_PASSWORD])
        smtp.send_message(msg)


# ------------------------------------------------------------------ logika ---
class HAMailer:
    """Obsługa jednego wpisu: wiele raportów, kolejka ponawiania, archiwum plików."""

    def __init__(self, hass: HomeAssistant, entry: ConfigEntry) -> None:
        self.hass = hass
        self.entry = entry
        self.stats: dict[str, dict] = {}
        self.pending: list[dict] = []  # niewysłane pliki czekające na ponowienie
        self.files: list[dict] = []  # indeks plików utworzonych przez dodatek
        self.next_runs: dict[str, datetime] = {}
        self._store = Store(hass, 1, f"{DOMAIN}.{entry.entry_id}")
        self._unsubs: dict[str, Callable] = {}
        self._retry_unsub: Callable | None = None
        self._cleanup_unsub: Callable | None = None
        self._lock = asyncio.Lock()

    # ---- dostęp do konfiguracji
    @property
    def cfg(self) -> dict:
        return {**self.entry.data, **self.entry.options}

    @property
    def reports(self) -> dict[str, dict]:
        return {r[CONF_ID]: r for r in self.cfg[CONF_REPORTS]}

    @property
    def folder(self) -> str:
        return self.hass.config.path(self.cfg.get(CONF_STORAGE_PATH) or "ha_mailer")

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

    # ---- stan trwały
    async def async_load(self) -> None:
        saved = await self._store.async_load() or {}
        ids = set(self.reports)
        for rid in ids:
            self.stats[rid] = {**DEFAULT_STATS, **saved.get("stats", {}).get(rid, {})}
        self.pending = [i for i in saved.get("pending", []) if i["report"] in ids]
        self.files = list(saved.get("files", []))

    async def _save(self) -> None:
        await self._store.async_save(
            {"stats": self.stats, "pending": self.pending, "files": self.files}
        )

    def _notify(self) -> None:
        async_dispatcher_send(self.hass, SIGNAL.format(self.entry.entry_id))

    # ---- harmonogramy
    def start(self) -> None:
        for rid in self.reports:
            self._schedule_report(rid)
        self._schedule_retry()
        self._cleanup_unsub = async_track_time_interval(
            self.hass, self._cleanup, timedelta(hours=1)
        )
        self.hass.async_create_task(self._cleanup())

    def stop(self) -> None:
        for rid in list(self._unsubs):
            self._cancel_report(rid)
        if self._retry_unsub:
            self._retry_unsub()
            self._retry_unsub = None
        if self._cleanup_unsub:
            self._cleanup_unsub()
            self._cleanup_unsub = None

    def _cancel_report(self, rid: str) -> None:
        unsub = self._unsubs.pop(rid, None)
        if unsub:
            unsub()

    def _schedule_report(self, rid: str) -> None:
        self._cancel_report(rid)
        report = self.reports.get(rid)
        if not report:
            return
        nxt = croniter(report[CONF_CRON], dt_util.now()).get_next(datetime)
        self.next_runs[rid] = nxt
        self._unsubs[rid] = async_track_point_in_time(
            self.hass, partial(self._fire_report, rid), nxt
        )
        self._notify()

    async def _fire_report(self, rid: str, _now: datetime) -> None:
        self._unsubs.pop(rid, None)
        try:
            await self.async_run_report(rid, queue_on_fail=True)
        except HomeAssistantError:
            pass  # błąd zarejestrowany (log, powiadomienie, encje), plik trafił do kolejki
        except Exception:  # noqa: BLE001
            _LOGGER.exception("Nieoczekiwany błąd harmonogramu HA_mailer (%s)", rid)
        finally:
            self._schedule_report(rid)

    def _schedule_retry(self) -> None:
        if self._retry_unsub:
            self._retry_unsub()
            self._retry_unsub = None
        times = [dt_util.parse_datetime(i["next_try"]) for i in self.pending if i.get("next_try")]
        times = [t for t in times if t]
        if times:
            self._retry_unsub = async_track_point_in_time(self.hass, self._retry_due, min(times))
        self._notify()

    # ---- zbieranie danych do raportu
    def _collect_rows(self, report: dict) -> list[tuple[bool, dict]]:
        """Kolejność = kolejność wierszy w konfiguracji raportu (nigdy się nie zmienia).

        Zwraca listę (czy_wiersz_z_samym_opisem, wartości_kolumn).
        """
        out: list[tuple[bool, dict]] = []
        number = 0
        for row in report[CONF_ROWS]:
            desc = row.get(ROW_DESC) or None
            entity_id = row.get(ROW_ENTITY)
            if not entity_id:
                out.append((True, {"opis": desc}))
                continue
            number += 1
            state = self.hass.states.get(entity_id)
            values = {"lp": number, "opis": desc, "encja": entity_id}
            if state is None:
                values.update(nazwa=entity_id, wartosc="brak encji", jednostka="")
            else:
                values.update(
                    nazwa=state.name,
                    wartosc=_to_value(state.state),
                    jednostka=state.attributes.get("unit_of_measurement", ""),
                    zmiana=dt_util.as_local(state.last_changed).replace(tzinfo=None, microsecond=0),
                )
            out.append((False, values))
        return out

    # ---- rejestr wyników
    async def _record(
        self, rid: str, now: datetime, error: str | None, size_kb: float | None = None, retried: bool = False
    ) -> None:
        st = self.stats.setdefault(rid, dict(DEFAULT_STATS))
        name = self.reports.get(rid, {}).get(CONF_NAME, rid)
        nid = f"{DOMAIN}_{self.entry.entry_id}_{rid}"
        st["last_attempt"] = now.isoformat()
        if error is None:
            st["ok"] += 1
            st["status"] = "ok"
            st["last_ok"] = now.isoformat()
            st["last_file_kb"] = size_kb
            if retried:
                st["retry_ok"] += 1
            if not self.pending_for(rid):
                persistent_notification.async_dismiss(self.hass, nid)
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
                    f"\n\nPlik został zapisany, a wysyłka będzie ponawiana co {minutes} min "
                    f"(oczekujące: {len(self.pending_for(rid))})."
                )
            persistent_notification.async_create(
                self.hass,
                f"Raport „{name}” - błąd wysyłki ({now.strftime(DT_FMT)}):\n\n{error}{retry_note}",
                title="HA_mailer - błąd wysyłki",
                notification_id=nid,
            )
        await self._save()
        self._notify()

    # ---- wysyłka
    async def _deliver(self, item: dict, payload: bytes) -> str | None:
        """Wysyła plik z `item`. Zwraca opis błędu albo None przy sukcesie."""
        report = self.reports[item["report"]]
        created = dt_util.parse_datetime(item["created"]) or dt_util.now()
        created_txt = dt_util.as_local(created).strftime(DT_FMT)
        if item["attempts"] == 0:
            body = (
                f"Raport „{report[CONF_NAME]}” z {created_txt}. "
                f"W załączniku wartości {item['entities']} encji."
            )
        else:
            body = (
                f"Raport „{report[CONF_NAME]}” z {created_txt}. Wysyłka była opóźniona "
                f"(nieudane próby: {item['attempts']}), ponowiona {dt_util.now().strftime(DT_FMT)}. "
                f"W załączniku wartości {item['entities']} encji z chwili utworzenia raportu."
            )
        try:
            await self.hass.async_add_executor_job(
                _send_mail, self.cfg, report[CONF_RECIPIENT], item["file"], item["subject"], payload, body
            )
        except smtplib.SMTPAuthenticationError:
            return "Gmail odrzucił logowanie - sprawdź adres konta i hasło aplikacji."
        except smtplib.SMTPRecipientsRefused:
            return "Serwer odrzucił adres odbiorcy - sprawdź adres."
        except (OSError, smtplib.SMTPException) as err:
            return f"Błąd połączenia/SMTP: {err}"
        except Exception as err:  # noqa: BLE001
            return f"Nieoczekiwany błąd: {err!r}"
        return None

    async def async_run_report(self, rid: str, *, queue_on_fail: bool) -> None:
        """Utwórz plik, zapisz go na dysku i wyślij. Przy błędzie rzuca HomeAssistantError.

        queue_on_fail=True (harmonogram): nieudany plik trafia do kolejki ponawiania.
        """
        report = self.reports[rid]
        async with self._lock:
            now = dt_util.now()
            rows = self._collect_rows(report)
            name = f"HA_mailer_{rid}_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            folder = self.folder
            item = {
                "id": uuid.uuid4().hex,
                "report": rid,
                "dir": folder,
                "file": name,
                "subject": f"{report[CONF_SUBJECT]} {now.strftime(DT_FMT)}",
                "created": now.isoformat(),
                "entities": sum(1 for is_text, _ in rows if not is_text),
                "attempts": 0,
                "last_error": None,
                "next_try": None,
            }
            try:
                payload = await self.hass.async_add_executor_job(
                    _build_xlsx, list(report[CONF_COLUMNS]), rows, now.replace(tzinfo=None, microsecond=0)
                )
                await self.hass.async_add_executor_job(_write_file, folder, name, payload)
            except Exception as err:  # noqa: BLE001
                error = f"Nie można utworzyć lub zapisać pliku w {folder}: {err}"
                await self._record(rid, now, error)
                raise HomeAssistantError(f"HA_mailer: {error}") from err
            self.files.append({"dir": folder, "file": name, "report": rid, "created": item["created"]})

            error = await self._deliver(item, payload)
            if error is None:
                await self._record(rid, now, None, round(len(payload) / 1024, 1))
                return
            if queue_on_fail:
                item["attempts"] = 1
                item["last_error"] = error
                item["next_try"] = (now + self.retry_interval).isoformat()
                self.pending.append(item)
                self._schedule_retry()
            await self._record(rid, now, error)
            raise HomeAssistantError(f"HA_mailer: {error}")

    async def _retry_due(self, _now: datetime) -> None:
        self._retry_unsub = None
        try:
            async with self._lock:
                limit = dt_util.now() + timedelta(seconds=1)
                due = [
                    i
                    for i in self.pending
                    if (dt_util.parse_datetime(i["next_try"]) or limit) <= limit
                ]
                for item in due:
                    if item in self.pending:
                        await self._retry_item(item)
        except Exception:  # noqa: BLE001
            _LOGGER.exception("Nieoczekiwany błąd ponawiania wysyłki HA_mailer")
        finally:
            self._schedule_retry()

    async def _retry_item(self, item: dict) -> None:
        now = dt_util.now()
        rid = item["report"]
        if rid not in self.reports:  # raport usunięty z konfiguracji
            self.pending.remove(item)
            await self._save()
            return
        try:
            payload = await self.hass.async_add_executor_job(
                _read_file, os.path.join(item["dir"], item["file"])
            )
        except OSError as err:
            self.pending.remove(item)
            await self._record(
                rid, now, f"Plik {item['file']} jest niedostępny ({err}) - porzucono ponawianie."
            )
            return
        error = await self._deliver(item, payload)
        if error is None:
            self.pending.remove(item)
            await self._record(rid, now, None, round(len(payload) / 1024, 1), retried=True)
        else:
            item["attempts"] += 1
            item["last_error"] = error
            item["next_try"] = (now + self.retry_interval).isoformat()
            await self._record(rid, now, error)

    async def async_clear_pending(self, rids: list[str] | None = None) -> int:
        """Porzuć oczekujące ponowienia (pliki zostają na dysku do końca retencji)."""
        async with self._lock:
            before = len(self.pending)
            self.pending = [i for i in self.pending if rids is not None and i["report"] not in rids]
            for rid in rids if rids is not None else list(self.reports):
                if not self.pending_for(rid):
                    nid = f"{DOMAIN}_{self.entry.entry_id}_{rid}"
                    persistent_notification.async_dismiss(self.hass, nid)
            await self._save()
            self._schedule_retry()
            return before - len(self.pending)

    # ---- archiwum plików
    async def _cleanup(self, _now: datetime | None = None) -> None:
        """Usuń pliki starsze niż retencja (poza czekającymi na ponowienie)."""
        days = int(self.cfg.get(CONF_RETENTION, DEFAULT_RETENTION))
        if days <= 0:
            return
        async with self._lock:
            cutoff = dt_util.now() - timedelta(days=days)
            busy = {(i["dir"], i["file"]) for i in self.pending}
            keep: list[dict] = []
            remove: list[tuple[str, str]] = []
            for f in self.files:
                created = dt_util.parse_datetime(f["created"])
                if (f["dir"], f["file"]) in busy or created is None or created > cutoff:
                    keep.append(f)
                else:
                    remove.append((f["dir"], f["file"]))
            if remove:
                await self.hass.async_add_executor_job(_delete_files, remove)
                self.files = keep
                await self._save()
                _LOGGER.debug("HA_mailer: usunięto %s starych plików", len(remove))
                self._notify()


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
        await store.async_save({"stats": {"raport": old}, "pending": [], "files": []})

    @callback
    def _migrate_unique_id(ent):
        suffix = ent.unique_id.removeprefix(f"{entry.entry_id}_")
        if suffix in ("status", "last_sent", "stats"):
            return {"new_unique_id": f"{entry.entry_id}_raport_{suffix}"}
        return None

    await er.async_migrate_entries(hass, entry.entry_id, _migrate_unique_id)
    hass.config_entries.async_update_entry(entry, data=new, options={}, version=3)
    return True


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    mailer = HAMailer(hass, entry)
    await mailer.async_load()
    hass.data.setdefault(DOMAIN, {})[entry.entry_id] = mailer
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
                    await mailer.async_run_report(rid, queue_on_fail=False)
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
    ok = await hass.config_entries.async_unload_platforms(entry, PLATFORMS)
    if ok:
        hass.data[DOMAIN].pop(entry.entry_id, None)
        if not hass.data[DOMAIN]:
            hass.services.async_remove(DOMAIN, "send_now")
            hass.services.async_remove(DOMAIN, "clear_pending")
    return ok


async def async_remove_entry(hass: HomeAssistant, entry: ConfigEntry) -> None:
    await Store(hass, 1, f"{DOMAIN}.{entry.entry_id}").async_remove()
