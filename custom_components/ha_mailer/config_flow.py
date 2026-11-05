"""Config flow (UI + import z YAML) i opcje HA_mailer."""
from __future__ import annotations

import json
import os
import smtplib
import ssl
from typing import Any

import voluptuous as vol
import yaml
from croniter import croniter

from homeassistant.config_entries import ConfigEntry, ConfigFlow, ConfigFlowResult, OptionsFlow
from homeassistant.core import callback
from homeassistant.helpers import selector
from homeassistant.util import slugify

from . import (
    COLUMN_LABELS,
    CONF_BCC,
    CONF_BODY,
    CONF_CC,
    CONF_COLUMNS,
    CONF_CRON,
    CONF_ENTITIES,
    CONF_GMAIL_PASSWORD,
    CONF_GMAIL_USER,
    CONF_ID,
    CONF_NAME,
    CONF_RECIPIENT,
    CONF_REPORTS,
    CONF_RETENTION,
    CONF_RETRY,
    CONF_ROWS,
    CONF_STORAGE_PATH,
    CONF_SUBJECT,
    CONF_YAML_HASH,
    DEFAULT_COLUMNS,
    DEFAULT_CRON,
    DEFAULT_FOLDER,
    DEFAULT_RETENTION,
    DEFAULT_RETRY,
    DEFAULT_SUBJECT,
    DOMAIN,
    EMAIL_RE,
    ERR_OPIS,
    GMAIL_HOST,
    GMAIL_PORT,
    LAYOUT_SCHEMA,
    ROW_ATTR,
    ROW_DESC,
    ROW_ENTITY,
    ROW_SHEET,
    is_text_row,
    layout_error,
    split_recipients,
    yaml_digest,
)

ERROR_TEXTS = {
    "no_name": "brak nazwy",
    "invalid_cron": "nieprawidłowy harmonogram",
    "invalid_recipient": "nieprawidłowy adres odbiorcy",
    "invalid_cc": "nieprawidłowy adres DW",
    "invalid_bcc": "nieprawidłowy adres UDW",
    "no_subject": "pusty tytuł",
    "no_columns": "brak kolumn",
    "opis_required": "wiersze bez encji wymagają kolumny „Opis”",
    "invalid_layout": "niepoprawny układ",
}


# ------------------------------------------------------------ pomocnicze ---
def _plain(value: Any) -> Any:
    """Kopia złożona wyłącznie z typów JSON (bezpieczna do YAML i do wpisu)."""
    return json.loads(json.dumps(value, default=str))


def _try_login(user: str, password: str) -> None:
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=ctx, timeout=20) as smtp:
        smtp.login(user, password)


def _check_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)
    probe = os.path.join(path, ".ha_mailer_test")
    with open(probe, "w", encoding="utf-8") as fh:
        fh.write("ok")
    os.remove(probe)


def _normalize(data: dict[str, Any]) -> dict[str, Any]:
    out = dict(data)
    for key in (CONF_GMAIL_USER, CONF_RECIPIENT, CONF_CC, CONF_BCC, CONF_SUBJECT, CONF_NAME, CONF_STORAGE_PATH):
        if isinstance(out.get(key), str):
            out[key] = out[key].strip()
    if isinstance(out.get(CONF_BODY), str):
        out[CONF_BODY] = out[CONF_BODY].strip()
    if isinstance(out.get(CONF_CRON), str):
        out[CONF_CRON] = " ".join(out[CONF_CRON].split())
    if isinstance(out.get(CONF_GMAIL_PASSWORD), str):
        out[CONF_GMAIL_PASSWORD] = out[CONF_GMAIL_PASSWORD].replace(" ", "")
    for key in (CONF_RETRY, CONF_RETENTION):
        if key in out:
            out[key] = int(out[key])
    for key in (CONF_CC, CONF_BCC, CONF_BODY):  # puste pola opcjonalne nie trafiają do wpisu
        if key in out and not out[key]:
            out.pop(key)
    return out


def _make_id(name: str, existing: list[str]) -> str:
    base = slugify(name) or "raport"
    rid, n = base, 2
    while rid in existing:
        rid, n = f"{base}_{n}", n + 1
    return rid


async def _validate_global(hass, d: dict[str, Any]) -> dict[str, str]:
    errors: dict[str, str] = {}
    if not EMAIL_RE.match(d.get(CONF_GMAIL_USER, "")):
        errors[CONF_GMAIL_USER] = "invalid_email"
    if int(d.get(CONF_RETRY, 0)) < 1:
        errors[CONF_RETRY] = "invalid_retry"
    path = d.get(CONF_STORAGE_PATH, "")
    if not path:
        errors[CONF_STORAGE_PATH] = "path_not_writable"
    else:
        try:
            await hass.async_add_executor_job(_check_dir, hass.config.path(path))
        except OSError:
            errors[CONF_STORAGE_PATH] = "path_not_writable"
    if errors:
        return errors
    try:
        await hass.async_add_executor_job(_try_login, d[CONF_GMAIL_USER], d[CONF_GMAIL_PASSWORD])
    except smtplib.SMTPAuthenticationError:
        errors["base"] = "auth_failed"
    except (OSError, smtplib.SMTPException):
        errors["base"] = "cannot_connect"
    return errors


def _valid_addresses(text: str | None, required: bool) -> bool:
    items = split_recipients(text)
    if not items:
        return not required
    return all(EMAIL_RE.match(i) for i in items)


def _validate_report(d: dict[str, Any]) -> tuple[dict[str, str], str]:
    """Zwraca (błędy, szczegóły) - szczegóły trafiają do komunikatu {details}."""
    errors: dict[str, str] = {}
    details = ""
    if not d.get(CONF_NAME):
        errors[CONF_NAME] = "no_name"
    if not croniter.is_valid(d.get(CONF_CRON, "")):
        errors[CONF_CRON] = "invalid_cron"
    if not _valid_addresses(d.get(CONF_RECIPIENT), True):
        errors[CONF_RECIPIENT] = "invalid_recipient"
    if not _valid_addresses(d.get(CONF_CC), False):
        errors[CONF_CC] = "invalid_cc"
    if not _valid_addresses(d.get(CONF_BCC), False):
        errors[CONF_BCC] = "invalid_bcc"
    if not d.get(CONF_SUBJECT):
        errors[CONF_SUBJECT] = "no_subject"
    if not d.get(CONF_COLUMNS):
        errors[CONF_COLUMNS] = "no_columns"
    else:
        err = layout_error(d.get(CONF_COLUMNS), d.get(CONF_ROWS))
        if err:
            if ERR_OPIS in err:
                errors[CONF_COLUMNS] = "opis_required"
            else:
                errors["base"] = "invalid_layout"
                details = err
    return errors, details


def _global_schema(defaults: dict[str, Any], password_required: bool) -> vol.Schema:
    pw = vol.Required(CONF_GMAIL_PASSWORD) if password_required else vol.Optional(CONF_GMAIL_PASSWORD)
    return vol.Schema(
        {
            vol.Required(CONF_GMAIL_USER, default=defaults.get(CONF_GMAIL_USER, "")): selector.TextSelector(
                selector.TextSelectorConfig(type=selector.TextSelectorType.EMAIL)
            ),
            pw: selector.TextSelector(
                selector.TextSelectorConfig(type=selector.TextSelectorType.PASSWORD)
            ),
            vol.Required(CONF_RETRY, default=defaults.get(CONF_RETRY, DEFAULT_RETRY)): selector.NumberSelector(
                selector.NumberSelectorConfig(
                    min=1, max=1440, step=1, mode=selector.NumberSelectorMode.BOX,
                    unit_of_measurement="min",
                )
            ),
            vol.Required(
                CONF_RETENTION, default=defaults.get(CONF_RETENTION, DEFAULT_RETENTION)
            ): selector.NumberSelector(
                selector.NumberSelectorConfig(
                    min=0, max=3650, step=1, mode=selector.NumberSelectorMode.BOX,
                    unit_of_measurement="dni",
                )
            ),
            vol.Required(
                CONF_STORAGE_PATH, default=defaults.get(CONF_STORAGE_PATH) or DEFAULT_FOLDER
            ): selector.TextSelector(),
        }
    )


def _report_schema(defaults: dict[str, Any], with_entities: bool) -> vol.Schema:
    def optional(key: str) -> vol.Optional:
        return vol.Optional(key, description={"suggested_value": defaults.get(key)})

    fields: dict = {
        vol.Required(CONF_NAME, default=defaults.get(CONF_NAME, "")): selector.TextSelector(),
        vol.Required(CONF_CRON, default=defaults.get(CONF_CRON, DEFAULT_CRON)): selector.TextSelector(),
        vol.Required(CONF_RECIPIENT, default=defaults.get(CONF_RECIPIENT, "")): selector.TextSelector(),
        optional(CONF_CC): selector.TextSelector(),
        optional(CONF_BCC): selector.TextSelector(),
        vol.Required(CONF_SUBJECT, default=defaults.get(CONF_SUBJECT, DEFAULT_SUBJECT)): selector.TextSelector(),
        optional(CONF_BODY): selector.TextSelector(selector.TextSelectorConfig(multiline=True)),
        vol.Required(CONF_COLUMNS, default=defaults.get(CONF_COLUMNS, DEFAULT_COLUMNS)): selector.SelectSelector(
            selector.SelectSelectorConfig(
                options=[{"value": k, "label": v} for k, v in COLUMN_LABELS.items()],
                multiple=True,
                mode=selector.SelectSelectorMode.DROPDOWN,
                sort=False,
            )
        ),
    }
    if with_entities:
        fields[vol.Required(CONF_ENTITIES, default=defaults.get(CONF_ENTITIES, []))] = (
            selector.EntitySelector(selector.EntitySelectorConfig(multiple=True))
        )
    return vol.Schema(fields)


def _report_from_form(user_input: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    data = _normalize(user_input)
    entities = data.pop(CONF_ENTITIES, [])
    data[CONF_ROWS] = [{ROW_ENTITY: e} for e in entities]
    return data, entities


# ------------------------------------------------------------ config flow ---
class HAMailerConfigFlow(ConfigFlow, domain=DOMAIN):
    VERSION = 3
    _data: dict[str, Any]

    @staticmethod
    @callback
    def async_get_options_flow(config_entry: ConfigEntry) -> OptionsFlow:
        return HAMailerOptionsFlow()

    async def async_step_user(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        """Krok 1: wspólne ustawienia (Gmail, ponawianie, archiwum)."""
        errors: dict[str, str] = {}
        if user_input is not None:
            user_input = _normalize(user_input)
            errors = await _validate_global(self.hass, user_input)
            if not errors:
                self._data = user_input
                return await self.async_step_first_report()
        defaults = {k: v for k, v in (user_input or {}).items() if k != CONF_GMAIL_PASSWORD}
        return self.async_show_form(
            step_id="user", data_schema=_global_schema(defaults, True), errors=errors
        )

    async def async_step_first_report(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        """Krok 2: pierwszy raport (kolejne dodasz w Konfiguruj)."""
        errors: dict[str, str] = {}
        details = ""
        if user_input is not None:
            report, entities = _report_from_form(user_input)
            if not entities:
                errors[CONF_ENTITIES] = "no_entities"
            else:
                errors, details = _validate_report(report)
            if not errors:
                report[CONF_ID] = _make_id(report[CONF_NAME], [])
                data = {**self._data, CONF_REPORTS: [report]}
                return self.async_create_entry(title="HA_mailer", data=data)
            user_input = {**report, CONF_ENTITIES: entities}
        return self.async_show_form(
            step_id="first_report",
            data_schema=_report_schema(user_input or {}, True),
            errors=errors,
            description_placeholders={"details": details},
        )

    async def async_step_import(self, import_data: dict[str, Any]) -> ConfigFlowResult:
        """Import z configuration.yaml. Aktualizuje wpis tylko, gdy YAML się zmienił."""
        conf = _plain(import_data)
        uid = f"yaml:{conf[CONF_ID]}"
        digest = yaml_digest(conf)
        data = {**conf, CONF_YAML_HASH: digest}
        await self.async_set_unique_id(uid)
        for entry in self._async_current_entries():
            if entry.unique_id == uid:
                if entry.data.get(CONF_YAML_HASH) == digest:
                    return self.async_abort(reason="already_configured")
                # YAML się zmienił: nadpisuje wpis, w tym wcześniejsze zmiany z UI
                self.hass.config_entries.async_update_entry(entry, data=data, options={})
                return self.async_abort(reason="yaml_updated")
        return self.async_create_entry(title="HA_mailer (YAML)", data=data)


# ----------------------------------------------------------- opcje (menu) ---
class HAMailerOptionsFlow(OptionsFlow):
    """Wszystkie zmiany trafiają od razu do kopii roboczej. Trwale zapisuje je tylko
    „Zapisz wszystko i zamknij”, a „Odrzuć zmiany i zamknij” kasuje całą kopię.
    """

    def __init__(self) -> None:
        self._opts: dict[str, Any] = {}
        self._rid = ""  # aktualnie edytowany raport
        self._idx = 0
        self._warning = ""

    # ---- pomocnicze
    def _report(self) -> dict[str, Any]:
        for r in self._opts[CONF_REPORTS]:
            if r[CONF_ID] == self._rid:
                return r
        raise KeyError(self._rid)

    @property
    def _rows(self) -> list[dict[str, str]]:
        return self._report().setdefault(CONF_ROWS, [])

    def _take_warning(self) -> str:
        text = f"**{self._warning}**\n\n" if self._warning else ""
        self._warning = ""
        return text

    def _reports_summary(self) -> str:
        reports = self._opts.get(CONF_REPORTS, [])
        if not reports:
            return "(brak raportów)"
        return "\n".join(
            f"{i}. **{r[CONF_NAME]}** - `{r[CONF_CRON]}` → {r[CONF_RECIPIENT]} ({len(r[CONF_ROWS])} wierszy)"
            for i, r in enumerate(reports, start=1)
        )

    def _report_select(self, multiple: bool = False) -> selector.SelectSelector:
        return selector.SelectSelector(
            selector.SelectSelectorConfig(
                options=[
                    {"value": r[CONF_ID], "label": r[CONF_NAME]} for r in self._opts[CONF_REPORTS]
                ],
                multiple=multiple,
                mode=selector.SelectSelectorMode.LIST,
            )
        )

    @staticmethod
    def _label(i: int, row: dict[str, str], md: bool = False) -> str:
        if row.get(ROW_SHEET):
            return f"{i}. [arkusz] {row[ROW_SHEET]}"
        ent, attr, desc = row.get(ROW_ENTITY), row.get(ROW_ATTR), row.get(ROW_DESC)
        if ent:
            base = f"`{ent}`" if md else ent
            if attr:
                base += f" ({attr})"
            return f"{i}. {base}" + (f" - {desc}" if desc else "")
        return f"{i}. [opis] {desc}"

    def _rows_summary(self) -> str:
        rows = self._rows
        if not rows:
            return "(brak wierszy)"
        return "\n".join(self._label(i, r, True) for i, r in enumerate(rows, start=1))

    def _row_select(self, multiple: bool = False) -> selector.SelectSelector:
        return selector.SelectSelector(
            selector.SelectSelectorConfig(
                options=[
                    {"value": str(i), "label": self._label(i + 1, r)} for i, r in enumerate(self._rows)
                ],
                multiple=multiple,
                mode=selector.SelectSelectorMode.LIST,
            )
        )

    @staticmethod
    def _position(maximum: int) -> selector.NumberSelector:
        return selector.NumberSelector(
            selector.NumberSelectorConfig(
                min=1, max=max(maximum, 1), step=1, mode=selector.NumberSelectorMode.BOX
            )
        )

    def _insert(self, row: dict[str, str], position: Any) -> None:
        rows = self._rows
        idx = len(rows) if not position else max(0, min(int(position) - 1, len(rows)))
        rows.insert(idx, row)

    def _ensure_opis(self) -> None:
        """Wiersze z samym opisem wymagają kolumny „Opis” - dodaj ją, jeśli brakuje."""
        report = self._report()
        columns = report.setdefault(CONF_COLUMNS, list(DEFAULT_COLUMNS))
        if "opis" not in columns and any(is_text_row(r) for r in report[CONF_ROWS]):
            columns.insert(1 if columns and columns[0] == "lp" else 0, "opis")
            self._warning = "Dodano kolumnę „Opis”, bo raport ma wiersze bez encji."

    # ---- hub
    def _hub(self) -> ConfigFlowResult:
        return self.async_show_menu(
            step_id="init",
            menu_options=["settings", "report_add", "report_edit", "report_delete", "save", "discard"],
            description_placeholders={
                "reports": self._reports_summary(),
                "warning": self._take_warning(),
            },
        )

    async def async_step_init(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        self._opts = _plain({**self.config_entry.data, **self.config_entry.options})
        self._opts.setdefault(CONF_REPORTS, [])
        self._warning = ""
        return self._hub()

    async def async_step_save(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        """Jedyne miejsce, w którym zmiany są trwale zapisywane."""
        if not self._opts[CONF_REPORTS]:
            self._warning = "Dodaj co najmniej jeden raport."
            return self._hub()
        for report in self._opts[CONF_REPORTS]:
            errors, details = _validate_report(report)
            if errors:
                reason = details or ", ".join(ERROR_TEXTS.get(e, e) for e in errors.values())
                self._warning = f"Nie można zapisać. Raport „{report[CONF_NAME]}”: {reason}."
                return self._hub()
        return self.async_create_entry(title="", data=self._opts)

    async def async_step_discard(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        return self.async_abort(reason="discarded")

    # ---- ustawienia wspólne
    async def async_step_settings(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        if user_input is not None:
            user_input = _normalize(user_input)
            if not user_input.get(CONF_GMAIL_PASSWORD):
                user_input[CONF_GMAIL_PASSWORD] = self._opts[CONF_GMAIL_PASSWORD]
            errors = await _validate_global(self.hass, {**self._opts, **user_input})
            if not errors:
                self._opts.update(user_input)
                return self._hub()
        defaults = {k: v for k, v in (user_input or self._opts).items() if k != CONF_GMAIL_PASSWORD}
        return self.async_show_form(
            step_id="settings", data_schema=_global_schema(defaults, False), errors=errors
        )

    # ---- raporty: dodawanie, wybór, usuwanie
    async def async_step_report_add(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        details = ""
        if user_input is not None:
            report, entities = _report_from_form(user_input)
            if not entities:
                errors[CONF_ENTITIES] = "no_entities"
            else:
                errors, details = _validate_report(report)
            if not errors:
                ids = [r[CONF_ID] for r in self._opts[CONF_REPORTS]]
                report[CONF_ID] = _make_id(report[CONF_NAME], ids)
                self._opts[CONF_REPORTS].append(report)
                self._rid = report[CONF_ID]
                return self._report_menu()
            user_input = {**report, CONF_ENTITIES: entities}
        return self.async_show_form(
            step_id="report_add",
            data_schema=_report_schema(user_input or {}, True),
            errors=errors,
            description_placeholders={"details": details},
        )

    async def async_step_report_edit(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if not self._opts[CONF_REPORTS]:
            self._warning = "Brak raportów."
            return self._hub()
        if user_input is not None:
            self._rid = user_input["report"]
            return self._report_menu()
        return self.async_show_form(
            step_id="report_edit",
            data_schema=vol.Schema({vol.Required("report"): self._report_select()}),
        )

    async def async_step_report_delete(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if not self._opts[CONF_REPORTS]:
            self._warning = "Brak raportów."
            return self._hub()
        if user_input is not None:
            drop = set(user_input.get("reports", []))
            self._opts[CONF_REPORTS] = [r for r in self._opts[CONF_REPORTS] if r[CONF_ID] not in drop]
            return self._hub()
        return self.async_show_form(
            step_id="report_delete",
            data_schema=vol.Schema({vol.Optional("reports"): self._report_select(multiple=True)}),
        )

    # ---- menu jednego raportu
    def _report_menu(self) -> ConfigFlowResult:
        r = self._report()
        return self.async_show_menu(
            step_id="report_menu",
            menu_options=["report_settings", "rows_visual", "layout_text", "report_back", "save"],
            description_placeholders={
                "name": r[CONF_NAME],
                "info": f"`{r[CONF_CRON]}` → {r[CONF_RECIPIENT]}, {len(r[CONF_ROWS])} wierszy",
                "warning": self._take_warning(),
            },
        )

    async def async_step_report_back(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        return self._hub()

    async def async_step_report_settings(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        details = ""
        report = self._report()
        if user_input is not None:
            user_input = _normalize(user_input)
            candidate = {k: v for k, v in report.items() if k not in (CONF_CC, CONF_BCC, CONF_BODY)}
            candidate.update(user_input)
            errors, details = _validate_report(candidate)
            if not errors:
                report.clear()
                report.update(candidate)
                return self._report_menu()
            report = candidate
        return self.async_show_form(
            step_id="report_settings",
            data_schema=_report_schema(report, False),
            errors=errors,
            description_placeholders={"details": details},
        )

    # ---- edytor graficzny wierszy (zmiany od razu w kopii roboczej)
    async def async_step_rows_visual(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        return self.async_show_menu(
            step_id="rows_visual",
            menu_options=[
                "add_entity", "add_text", "add_sheet", "edit_row", "delete_row", "move_row", "rows_back",
            ],
            description_placeholders={
                "name": self._report()[CONF_NAME],
                "rows": self._rows_summary(),
                "warning": self._take_warning(),
            },
        )

    async def async_step_rows_back(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        return self._report_menu()

    async def async_step_add_entity(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if user_input is not None:
            row = {ROW_ENTITY: user_input[ROW_ENTITY]}
            attr = (user_input.get(ROW_ATTR) or "").strip()
            if attr:
                row[ROW_ATTR] = attr
            desc = (user_input.get(ROW_DESC) or "").strip()
            if desc:
                row[ROW_DESC] = desc
            self._insert(row, user_input.get("position"))
            return await self.async_step_rows_visual()
        return self.async_show_form(
            step_id="add_entity",
            data_schema=vol.Schema(
                {
                    vol.Required(ROW_ENTITY): selector.EntitySelector(),
                    vol.Optional(ROW_ATTR): selector.TextSelector(),
                    vol.Optional(ROW_DESC): selector.TextSelector(),
                    vol.Optional("position"): self._position(len(self._rows) + 1),
                }
            ),
        )

    async def async_step_add_text(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        if user_input is not None:
            desc = (user_input.get(ROW_DESC) or "").strip()
            if desc:
                self._insert({ROW_DESC: desc}, user_input.get("position"))
                self._ensure_opis()
                return await self.async_step_rows_visual()
            errors["base"] = "row_empty"
        return self.async_show_form(
            step_id="add_text",
            data_schema=vol.Schema(
                {
                    vol.Required(ROW_DESC): selector.TextSelector(),
                    vol.Optional("position"): self._position(len(self._rows) + 1),
                }
            ),
            errors=errors,
        )

    async def async_step_add_sheet(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        if user_input is not None:
            name = (user_input.get(ROW_SHEET) or "").strip()
            if name:
                self._insert({ROW_SHEET: name}, user_input.get("position"))
                return await self.async_step_rows_visual()
            errors["base"] = "row_empty"
        return self.async_show_form(
            step_id="add_sheet",
            data_schema=vol.Schema(
                {
                    vol.Required(ROW_SHEET): selector.TextSelector(),
                    vol.Optional("position"): self._position(len(self._rows) + 1),
                }
            ),
            errors=errors,
        )

    async def async_step_edit_row(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if not self._rows:
            self._warning = "Brak wierszy."
            return await self.async_step_rows_visual()
        if user_input is not None:
            self._idx = int(user_input["row"])
            return await self.async_step_edit_row_form()
        return self.async_show_form(
            step_id="edit_row",
            data_schema=vol.Schema({vol.Required("row"): self._row_select()}),
        )

    async def async_step_edit_row_form(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        row = self._rows[self._idx]
        is_sheet = bool(row.get(ROW_SHEET))
        if user_input is not None:
            new: dict[str, str] = {}
            if is_sheet:
                if (user_input.get(ROW_SHEET) or "").strip():
                    new[ROW_SHEET] = user_input[ROW_SHEET].strip()
            else:
                if user_input.get(ROW_ENTITY):
                    new[ROW_ENTITY] = user_input[ROW_ENTITY]
                    attr = (user_input.get(ROW_ATTR) or "").strip()
                    if attr:
                        new[ROW_ATTR] = attr
                desc = (user_input.get(ROW_DESC) or "").strip()
                if desc:
                    new[ROW_DESC] = desc
            if new:
                self._rows[self._idx] = new
                self._ensure_opis()
                return await self.async_step_rows_visual()
            errors["base"] = "row_empty"
            row = user_input

        def suggest(key: str) -> dict[str, Any]:
            return {"suggested_value": row.get(key)}

        if is_sheet:
            schema = vol.Schema(
                {vol.Required(ROW_SHEET, description=suggest(ROW_SHEET)): selector.TextSelector()}
            )
        else:
            schema = vol.Schema(
                {
                    vol.Optional(ROW_ENTITY, description=suggest(ROW_ENTITY)): selector.EntitySelector(),
                    vol.Optional(ROW_ATTR, description=suggest(ROW_ATTR)): selector.TextSelector(),
                    vol.Optional(ROW_DESC, description=suggest(ROW_DESC)): selector.TextSelector(),
                }
            )
        return self.async_show_form(
            step_id="edit_row_form", data_schema=schema, errors=errors,
            description_placeholders={"kind": "arkusza" if is_sheet else "wiersza"},
        )

    async def async_step_delete_row(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if not self._rows:
            self._warning = "Brak wierszy."
            return await self.async_step_rows_visual()
        if user_input is not None:
            drop = {int(i) for i in user_input.get("rows", [])}
            report = self._report()
            report[CONF_ROWS] = [r for i, r in enumerate(report[CONF_ROWS]) if i not in drop]
            return await self.async_step_rows_visual()
        return self.async_show_form(
            step_id="delete_row",
            data_schema=vol.Schema({vol.Optional("rows"): self._row_select(multiple=True)}),
        )

    async def async_step_move_row(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if not self._rows:
            self._warning = "Brak wierszy."
            return await self.async_step_rows_visual()
        if user_input is not None:
            rows = self._rows
            item = rows.pop(int(user_input["row"]))
            pos = max(1, min(int(user_input["position"]), len(rows) + 1))
            rows.insert(pos - 1, item)
            return await self.async_step_rows_visual()
        return self.async_show_form(
            step_id="move_row",
            data_schema=vol.Schema(
                {
                    vol.Required("row"): self._row_select(),
                    vol.Required("position"): self._position(len(self._rows)),
                }
            ),
        )

    # ---- edycja tekstowa (YAML) bieżącego raportu
    async def async_step_layout_text(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        details = ""
        report = self._report()
        text = yaml.safe_dump(
            {CONF_COLUMNS: report.get(CONF_COLUMNS, DEFAULT_COLUMNS), CONF_ROWS: report[CONF_ROWS]},
            allow_unicode=True,
            sort_keys=False,
        )
        if user_input is not None:
            text = user_input["layout_yaml"]
            try:
                parsed = yaml.safe_load(text)
            except yaml.YAMLError as err:
                errors["base"] = "invalid_yaml"
                details = str(err).replace("\n", " ")
            else:
                try:
                    valid = LAYOUT_SCHEMA(parsed)
                except vol.Invalid as err:
                    if ERR_OPIS in str(err):
                        errors["base"] = "opis_required"
                    else:
                        errors["base"] = "invalid_layout"
                        details = str(err)
                else:
                    report[CONF_COLUMNS] = _plain(valid[CONF_COLUMNS])
                    report[CONF_ROWS] = _plain(valid[CONF_ROWS])
                    return self._report_menu()
        return self.async_show_form(
            step_id="layout_text",
            data_schema=vol.Schema(
                {
                    vol.Required(
                        "layout_yaml", description={"suggested_value": text}
                    ): selector.TextSelector(selector.TextSelectorConfig(multiline=True))
                }
            ),
            errors=errors,
            description_placeholders={"details": details, "name": report[CONF_NAME]},
        )
