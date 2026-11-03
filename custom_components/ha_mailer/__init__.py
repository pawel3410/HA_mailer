"""HA_mailer - migawka wybranych encji w Excelu wysyłana mailem z Gmaila wg cron."""
from __future__ import annotations

import logging
import math
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
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant, ServiceCall
from homeassistant.exceptions import HomeAssistantError, ServiceValidationError
from homeassistant.helpers.dispatcher import async_dispatcher_send
from homeassistant.helpers.event import async_track_point_in_time
from homeassistant.helpers.storage import Store
from homeassistant.util import dt as dt_util

_LOGGER = logging.getLogger(__name__)

DOMAIN = "ha_mailer"
PLATFORMS = ["sensor"]

CONF_ENTITIES = "entities"
CONF_CRON = "cron"
CONF_GMAIL_USER = "gmail_user"
CONF_GMAIL_PASSWORD = "gmail_app_password"
CONF_RECIPIENT = "recipient"
CONF_SUBJECT = "subject"

DEFAULT_CRON = "0 8 * * *"
DEFAULT_SUBJECT = "HA_mailer - raport"
GMAIL_HOST = "smtp.gmail.com"
GMAIL_PORT = 465

SIGNAL = DOMAIN + "_update_{}"
DT_FMT = "%Y-%m-%d %H:%M:%S"
XLSX_MIME = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
HEADERS = ["Lp.", "Encja", "Nazwa", "Wartość", "Jednostka", "Ostatnia zmiana"]
FIRST_DATA_ROW = 4  # wiersz 1: czas migawki, wiersz 3: nagłówki, dane od wiersza 4

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


def _to_value(raw: str):
    try:
        number = float(raw)
    except (ValueError, TypeError):
        return raw
    return number if math.isfinite(number) else raw


def _build_xlsx(rows: list[list], snapshot: datetime) -> bytes:
    """Układ stały: wiersz N pliku = N-ta encja z konfiguracji."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Migawka"
    ws["A1"] = "Czas migawki:"
    ws["A1"].font = Font(bold=True)
    ws["B1"] = snapshot
    ws["B1"].number_format = "yyyy-mm-dd hh:mm:ss"
    ws["B1"].alignment = Alignment(horizontal="left")
    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
    for i, row in enumerate(rows):
        for col, value in enumerate(row, start=1):
            cell = ws.cell(row=FIRST_DATA_ROW + i, column=col, value=value)
            if col == 6 and value is not None:
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
    for col, header in enumerate(HEADERS, start=1):
        longest = max([len(header)] + [len(str(r[col - 1])) for r in rows if r[col - 1] is not None])
        ws.column_dimensions[get_column_letter(col)].width = min(longest + 3, 60)
    ws.freeze_panes = "A4"
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _send_mail(cfg: dict, filename: str, subject: str, payload: bytes, body: str) -> None:
    recipients = [r.strip() for r in cfg[CONF_RECIPIENT].replace(";", ",").split(",") if r.strip()]
    msg = EmailMessage()
    msg["From"] = cfg[CONF_GMAIL_USER]
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)
    msg.add_attachment(payload, maintype=XLSX_MIME[0], subtype=XLSX_MIME[1], filename=filename)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=context, timeout=30) as smtp:
        smtp.login(cfg[CONF_GMAIL_USER], cfg[CONF_GMAIL_PASSWORD])
        smtp.send_message(msg)


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

    def _collect_rows(self) -> list[list]:
        """Kolejność wierszy = kolejność encji w konfiguracji (nigdy się nie zmienia)."""
        rows: list[list] = []
        for i, entity_id in enumerate(self.cfg[CONF_ENTITIES], start=1):
            state = self.hass.states.get(entity_id)
            if state is None:
                rows.append([i, entity_id, entity_id, "brak encji", "", None])
                continue
            changed = dt_util.as_local(state.last_changed).replace(tzinfo=None, microsecond=0)
            rows.append(
                [
                    i,
                    entity_id,
                    state.name,
                    _to_value(state.state),
                    state.attributes.get("unit_of_measurement", ""),
                    changed,
                ]
            )
        return rows

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
            body = f"Raport HA_mailer z {stamp}. W załączniku wartości {len(rows)} encji."
            payload = await self.hass.async_add_executor_job(
                _build_xlsx, rows, now.replace(tzinfo=None, microsecond=0)
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
