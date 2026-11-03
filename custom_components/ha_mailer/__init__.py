"""HA_mailer - migawka wybranych encji w Excelu wysyłana mailem z Gmaila wg cron."""
from __future__ import annotations

import logging
import math
import re
import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from io import BytesIO

import voluptuous as vol
from croniter import croniter
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from homeassistant.components import persistent_notification
from homeassistant.config_entries import SOURCE_IMPORT, ConfigEntry
from homeassistant.core import HomeAssistant, ServiceCall
from homeassistant.exceptions import HomeAssistantError, ServiceValidationError
import homeassistant.helpers.config_validation as cv
from homeassistant.helpers.dispatcher import async_dispatcher_send
from homeassistant.helpers.event import async_track_point_in_time
from homeassistant.helpers.storage import Store
from homeassistant.util import dt as dt_util

_LOGGER = logging.getLogger(__name__)

DOMAIN = "ha_mailer"
PLATFORMS = ["sensor"]

CONF_ID = "id"
CONF_ENTITIES = "entities"  # tylko w kroku początkowym UI (potem zamieniane na wiersze)
CONF_CRON = "cron"
CONF_GMAIL_USER = "gmail_user"
CONF_GMAIL_PASSWORD = "gmail_app_password"
CONF_RECIPIENT = "recipient"
CONF_SUBJECT = "subject"
CONF_COLUMNS = "columns"
CONF_ROWS = "rows"
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
    "fail": 0,
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
        raise vol.Invalid("Kolumny nie mogą się powtarzać.")
    return values


def _check_layout(conf: dict) -> dict:
    if any(not r.get(ROW_ENTITY) for r in conf[CONF_ROWS]) and "opis" not in conf[CONF_COLUMNS]:
        raise vol.Invalid(ERR_OPIS)
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
    vol.Schema(
        {vol.Required(CONF_COLUMNS): COLUMNS_SCHEMA, vol.Required(CONF_ROWS): ROWS_SCHEMA}
    ),
    _check_layout,
)
YAML_SCHEMA = vol.All(
    vol.Schema(
        {
            vol.Optional(CONF_ID): cv.string,
            vol.Required(CONF_CRON): _cron,
            vol.Required(CONF_GMAIL_USER): _email,
            vol.Required(CONF_GMAIL_PASSWORD): cv.string,
            vol.Required(CONF_RECIPIENT): _recipients,
            vol.Optional(CONF_SUBJECT, default=DEFAULT_SUBJECT): cv.string,
            vol.Optional(CONF_COLUMNS, default=DEFAULT_COLUMNS): COLUMNS_SCHEMA,
            vol.Required(CONF_ROWS): ROWS_SCHEMA,
        }
    ),
    _check_layout,
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


def _send_mail(cfg: dict, filename: str, subject: str, payload: bytes, body: str) -> None:
    msg = EmailMessage()
    msg["From"] = cfg[CONF_GMAIL_USER]
    msg["To"] = ", ".join(split_recipients(cfg[CONF_RECIPIENT]))
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(payload, maintype=XLSX_MIME[0], subtype=XLSX_MIME[1], filename=filename)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=context, timeout=30) as smtp:
        smtp.login(cfg[CONF_GMAIL_USER], cfg[CONF_GMAIL_PASSWORD])
        smtp.send_message(msg)


# ------------------------------------------------------------------ logika ---
class HAMailer:
    """Logika jednego wpisu konfiguracji + statystyki (zapisywane na dysku)."""

    def __init__(self, hass: HomeAssistant, entry: ConfigEntry) -> None:
        self.hass = hass
        self.entry = entry
        self.stats: dict = dict(DEFAULT_STATS)
        self.next_run: datetime | None = None
        self._store = Store(hass, 1, f"{DOMAIN}.{entry.entry_id}")
        self._unsub = None

    @property
    def cfg(self) -> dict:
        return {**self.entry.data, **self.entry.options}

    async def async_load(self) -> None:
        saved = await self._store.async_load()
        if saved:
            self.stats.update(saved)

    def _notify_entities(self) -> None:
        async_dispatcher_send(self.hass, SIGNAL.format(self.entry.entry_id))

    def schedule(self) -> None:
        self.cancel()
        self.next_run = croniter(self.cfg[CONF_CRON], dt_util.now()).get_next(datetime)
        self._unsub = async_track_point_in_time(self.hass, self._fire, self.next_run)
        self._notify_entities()

    def cancel(self) -> None:
        if self._unsub:
            self._unsub()
            self._unsub = None

    async def _fire(self, _now: datetime) -> None:
        self._unsub = None
        try:
            await self.async_send()
        except HomeAssistantError:
            pass  # błąd jest już zarejestrowany (log, powiadomienie, encje)
        except Exception:  # noqa: BLE001
            _LOGGER.exception("Nieoczekiwany błąd harmonogramu HA_mailer")
        finally:
            self.schedule()

    def _collect_rows(self) -> list[tuple[bool, dict]]:
        """Kolejność = kolejność wierszy w konfiguracji (nigdy się nie zmienia).

        Zwraca listę (czy_wiersz_z_samym_opisem, wartości_kolumn).
        """
        out: list[tuple[bool, dict]] = []
        number = 0
        for row in self.cfg[CONF_ROWS]:
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

    async def _record(self, now: datetime, error: str | None, size_kb: float | None = None) -> None:
        self.stats["last_attempt"] = now.isoformat()
        nid = f"{DOMAIN}_{self.entry.entry_id}"
        if error is None:
            self.stats["ok"] += 1
            self.stats["status"] = "ok"
            self.stats["last_ok"] = now.isoformat()
            self.stats["last_file_kb"] = size_kb
            persistent_notification.async_dismiss(self.hass, nid)
        else:
            self.stats["fail"] += 1
            self.stats["status"] = "error"
            self.stats["last_error"] = error
            self.stats["last_error_time"] = now.isoformat()
            _LOGGER.error("HA_mailer: wysyłka nieudana: %s", error)
            persistent_notification.async_create(
                self.hass,
                f"Nie udało się wysłać raportu ({now.strftime(DT_FMT)}):\n\n{error}",
                title="HA_mailer - błąd wysyłki",
                notification_id=nid,
            )
        await self._store.async_save(self.stats)
        self._notify_entities()

    async def async_send(self) -> None:
        """Zrób migawkę, wygeneruj Excela i wyślij. Przy błędzie rzuca HomeAssistantError."""
        cfg = self.cfg
        now = dt_util.now()
        error: str | None = None
        size_kb = None
        try:
            rows = self._collect_rows()
            stamp = now.strftime(DT_FMT)
            subject = f"{cfg[CONF_SUBJECT]} {stamp}"
            filename = f"HA_mailer_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            count = sum(1 for is_text, _ in rows if not is_text)
            body = f"Raport HA_mailer z {stamp}. W załączniku wartości {count} encji."
            payload = await self.hass.async_add_executor_job(
                _build_xlsx, list(cfg[CONF_COLUMNS]), rows, now.replace(tzinfo=None, microsecond=0)
            )
            size_kb = round(len(payload) / 1024, 1)
            await self.hass.async_add_executor_job(_send_mail, cfg, filename, subject, payload, body)
        except smtplib.SMTPAuthenticationError:
            error = "Gmail odrzucił logowanie - sprawdź adres konta i hasło aplikacji."
        except smtplib.SMTPRecipientsRefused:
            error = "Serwer odrzucił adres odbiorcy - sprawdź adres."
        except (OSError, smtplib.SMTPException) as err:
            error = f"Błąd połączenia/SMTP: {err}"
        except Exception as err:  # noqa: BLE001
            error = f"Nieoczekiwany błąd: {err!r}"
        await self._record(now, error, size_kb)
        if error:
            raise HomeAssistantError(f"HA_mailer: {error}")
        _LOGGER.info("HA_mailer: wysłano raport do %s", cfg[CONF_RECIPIENT])


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
    """v1 (lista encji) -> v2 (kolumny + wiersze)."""
    if entry.version == 1:

        def convert(src) -> dict:
            d = dict(src)
            if CONF_ENTITIES in d:
                d[CONF_ROWS] = [{ROW_ENTITY: e} for e in d.pop(CONF_ENTITIES)]
            return d

        data = convert(entry.data)
        data.setdefault(CONF_COLUMNS, list(LEGACY_COLUMNS))
        hass.config_entries.async_update_entry(
            entry, data=data, options=convert(entry.options), version=2
        )
    return True


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    mailer = HAMailer(hass, entry)
    await mailer.async_load()
    hass.data.setdefault(DOMAIN, {})[entry.entry_id] = mailer
    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)
    mailer.schedule()
    entry.async_on_unload(mailer.cancel)
    entry.async_on_unload(entry.add_update_listener(_async_reload))

    if not hass.services.has_service(DOMAIN, "send_now"):

        async def _send_now(call: ServiceCall) -> None:
            mailers: dict[str, HAMailer] = hass.data.get(DOMAIN, {})
            entry_id = call.data.get("entry_id")
            if entry_id and entry_id not in mailers:
                raise ServiceValidationError(f"Nieznany entry_id: {entry_id}")
            for m in [mailers[entry_id]] if entry_id else list(mailers.values()):
                await m.async_send()

        hass.services.async_register(
            DOMAIN, "send_now", _send_now, schema=vol.Schema({vol.Optional("entry_id"): str})
        )
    return True


async def _async_reload(hass: HomeAssistant, entry: ConfigEntry) -> None:
    await hass.config_entries.async_reload(entry.entry_id)


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    ok = await hass.config_entries.async_unload_platforms(entry, PLATFORMS)
    if ok:
        hass.data[DOMAIN].pop(entry.entry_id, None)
        if not hass.data[DOMAIN]:
            hass.services.async_remove(DOMAIN, "send_now")
    return ok


async def async_remove_entry(hass: HomeAssistant, entry: ConfigEntry) -> None:
    await Store(hass, 1, f"{DOMAIN}.{entry.entry_id}").async_remove()
