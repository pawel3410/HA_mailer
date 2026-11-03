"""Config flow (UI + import z YAML) i opcje HA_mailer."""
from __future__ import annotations

import hashlib
import json
import smtplib
import ssl
from typing import Any

import voluptuous as vol
import yaml
from croniter import croniter

from homeassistant.config_entries import ConfigEntry, ConfigFlow, ConfigFlowResult, OptionsFlow
from homeassistant.core import callback
from homeassistant.helpers import selector

from . import (
    COLUMN_LABELS,
    CONF_COLUMNS,
    CONF_CRON,
    CONF_ENTITIES,
    CONF_GMAIL_PASSWORD,
    CONF_GMAIL_USER,
    CONF_ID,
    CONF_RECIPIENT,
    CONF_ROWS,
    CONF_SUBJECT,
    CONF_YAML_HASH,
    DEFAULT_COLUMNS,
    DEFAULT_CRON,
    DEFAULT_SUBJECT,
    DOMAIN,
    EMAIL_RE,
    ERR_OPIS,
    GMAIL_HOST,
    GMAIL_PORT,
    LAYOUT_SCHEMA,
    ROW_DESC,
    ROW_ENTITY,
    layout_error,
    split_recipients,
)


def _plain(value: Any) -> Any:
    """Kopia złożona wyłącznie z typów JSON (bezpieczna do YAML i do wpisu)."""
    return json.loads(json.dumps(value, default=str))


def _try_login(user: str, password: str) -> None:
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=ctx, timeout=20) as smtp:
        smtp.login(user, password)


def _normalize(data: dict[str, Any]) -> dict[str, Any]:
    out = dict(data)
    for key in (CONF_GMAIL_USER, CONF_RECIPIENT, CONF_SUBJECT):
        if isinstance(out.get(key), str):
            out[key] = out[key].strip()
    if isinstance(out.get(CONF_CRON), str):
        out[CONF_CRON] = " ".join(out[CONF_CRON].split())
    if isinstance(out.get(CONF_GMAIL_PASSWORD), str):
        out[CONF_GMAIL_PASSWORD] = out[CONF_GMAIL_PASSWORD].replace(" ", "")
    return out


async def _validate(hass, d: dict[str, Any]) -> tuple[dict[str, str], str]:
    """Zwraca (błędy, szczegóły) - szczegóły trafiają do komunikatu {details}."""
    errors: dict[str, str] = {}
    details = ""
    if not croniter.is_valid(d.get(CONF_CRON, "")):
        errors[CONF_CRON] = "invalid_cron"
    if not EMAIL_RE.match(d.get(CONF_GMAIL_USER, "")):
        errors[CONF_GMAIL_USER] = "invalid_email"
    rcpt = split_recipients(d.get(CONF_RECIPIENT, ""))
    if not rcpt or not all(EMAIL_RE.match(r) for r in rcpt):
        errors[CONF_RECIPIENT] = "invalid_recipient"
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
    if errors:
        return errors, details
    try:
        await hass.async_add_executor_job(_try_login, d[CONF_GMAIL_USER], d[CONF_GMAIL_PASSWORD])
    except smtplib.SMTPAuthenticationError:
        errors["base"] = "auth_failed"
    except (OSError, smtplib.SMTPException):
        errors["base"] = "cannot_connect"
    return errors, details


def _settings_schema(defaults: dict[str, Any], password_required: bool, with_entities: bool) -> vol.Schema:
    fields: dict = {}
    if with_entities:
        fields[vol.Required(CONF_ENTITIES, default=defaults.get(CONF_ENTITIES, []))] = (
            selector.EntitySelector(selector.EntitySelectorConfig(multiple=True))
        )
    fields[vol.Required(CONF_CRON, default=defaults.get(CONF_CRON, DEFAULT_CRON))] = (
        selector.TextSelector()
    )
    fields[vol.Required(CONF_GMAIL_USER, default=defaults.get(CONF_GMAIL_USER, ""))] = (
        selector.TextSelector(selector.TextSelectorConfig(type=selector.TextSelectorType.EMAIL))
    )
    pw = vol.Required(CONF_GMAIL_PASSWORD) if password_required else vol.Optional(CONF_GMAIL_PASSWORD)
    fields[pw] = selector.TextSelector(
        selector.TextSelectorConfig(type=selector.TextSelectorType.PASSWORD)
    )
    fields[vol.Required(CONF_RECIPIENT, default=defaults.get(CONF_RECIPIENT, ""))] = (
        selector.TextSelector()
    )
    fields[vol.Required(CONF_SUBJECT, default=defaults.get(CONF_SUBJECT, DEFAULT_SUBJECT))] = (
        selector.TextSelector()
    )
    fields[vol.Required(CONF_COLUMNS, default=defaults.get(CONF_COLUMNS, DEFAULT_COLUMNS))] = (
        selector.SelectSelector(
            selector.SelectSelectorConfig(
                options=[{"value": k, "label": v} for k, v in COLUMN_LABELS.items()],
                multiple=True,
                mode=selector.SelectSelectorMode.DROPDOWN,
                sort=False,
            )
        )
    )
    return vol.Schema(fields)


class HAMailerConfigFlow(ConfigFlow, domain=DOMAIN):
    VERSION = 2

    @staticmethod
    @callback
    def async_get_options_flow(config_entry: ConfigEntry) -> OptionsFlow:
        return HAMailerOptionsFlow()

    async def async_step_user(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        details = ""
        if user_input is not None:
            user_input = _normalize(user_input)
            entities = user_input.pop(CONF_ENTITIES, [])
            user_input[CONF_ROWS] = [{ROW_ENTITY: e} for e in entities]
            if not entities:
                errors[CONF_ENTITIES] = "no_entities"
            else:
                errors, details = await _validate(self.hass, user_input)
            if not errors:
                return self.async_create_entry(title="HA_mailer", data=user_input)
            user_input[CONF_ENTITIES] = entities
        defaults = {k: v for k, v in (user_input or {}).items() if k != CONF_GMAIL_PASSWORD}
        return self.async_show_form(
            step_id="user",
            data_schema=_settings_schema(defaults, True, True),
            errors=errors,
            description_placeholders={"details": details},
        )

    async def async_step_import(self, import_data: dict[str, Any]) -> ConfigFlowResult:
        """Import z configuration.yaml. Aktualizuje wpis tylko, gdy YAML się zmienił."""
        conf = _plain(import_data)
        uid = f"yaml:{conf[CONF_ID]}"
        digest = hashlib.sha256(json.dumps(conf, sort_keys=True).encode()).hexdigest()
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


class HAMailerOptionsFlow(OptionsFlow):
    """Menu: ustawienia / edytor graficzny wierszy / edycja tekstowa układu."""

    def __init__(self) -> None:
        self._opts: dict[str, Any] = {}
        self._rows: list[dict[str, str]] = []
        self._idx = 0
        self._warning = ""

    # ---- pomocnicze
    @staticmethod
    def _label(i: int, row: dict[str, str], md: bool = False) -> str:
        ent, desc = row.get(ROW_ENTITY), row.get(ROW_DESC)
        if ent:
            base = f"`{ent}`" if md else ent
            return f"{i}. {base}" + (f" - {desc}" if desc else "")
        return f"{i}. [opis] {desc}"

    def _summary(self) -> str:
        if not self._rows:
            return "(brak wierszy)"
        return "\n".join(self._label(i, r, True) for i, r in enumerate(self._rows, start=1))

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

    def _position(self, maximum: int) -> selector.NumberSelector:
        return selector.NumberSelector(
            selector.NumberSelectorConfig(
                min=1, max=max(maximum, 1), step=1, mode=selector.NumberSelectorMode.BOX
            )
        )

    def _insert(self, row: dict[str, str], position: Any) -> None:
        idx = len(self._rows) if not position else max(0, min(int(position) - 1, len(self._rows)))
        self._rows.insert(idx, row)

    # ---- menu główne
    async def async_step_init(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        self._opts = _plain({**self.config_entry.data, **self.config_entry.options})
        self._rows = _plain(self._opts.get(CONF_ROWS, []))
        self._warning = ""
        return self.async_show_menu(
            step_id="init", menu_options=["settings", "rows_visual", "layout_text"]
        )

    # ---- ustawienia
    async def async_step_settings(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        details = ""
        if user_input is not None:
            user_input = _normalize(user_input)
            if not user_input.get(CONF_GMAIL_PASSWORD):
                user_input[CONF_GMAIL_PASSWORD] = self._opts[CONF_GMAIL_PASSWORD]
            candidate = {**self._opts, **user_input}
            errors, details = await _validate(self.hass, candidate)
            if not errors:
                return self.async_create_entry(title="", data=candidate)
        defaults = {k: v for k, v in (user_input or self._opts).items() if k != CONF_GMAIL_PASSWORD}
        return self.async_show_form(
            step_id="settings",
            data_schema=_settings_schema(defaults, False, False),
            errors=errors,
            description_placeholders={"details": details},
        )

    # ---- edytor graficzny wierszy
    async def async_step_rows_visual(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        warning = f"**{self._warning}**\n\n" if self._warning else ""
        self._warning = ""
        return self.async_show_menu(
            step_id="rows_visual",
            menu_options=["add_entity", "add_text", "edit_row", "delete_row", "move_row", "rows_save"],
            description_placeholders={"rows": self._summary(), "warning": warning},
        )

    async def async_step_add_entity(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if user_input is not None:
            row = {ROW_ENTITY: user_input[ROW_ENTITY]}
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
        if user_input is not None:
            new: dict[str, str] = {}
            if user_input.get(ROW_ENTITY):
                new[ROW_ENTITY] = user_input[ROW_ENTITY]
            desc = (user_input.get(ROW_DESC) or "").strip()
            if desc:
                new[ROW_DESC] = desc
            if new:
                self._rows[self._idx] = new
                return await self.async_step_rows_visual()
            errors["base"] = "row_empty"
            row = user_input
        return self.async_show_form(
            step_id="edit_row_form",
            data_schema=vol.Schema(
                {
                    vol.Optional(
                        ROW_ENTITY, description={"suggested_value": row.get(ROW_ENTITY)}
                    ): selector.EntitySelector(),
                    vol.Optional(
                        ROW_DESC, description={"suggested_value": row.get(ROW_DESC)}
                    ): selector.TextSelector(),
                }
            ),
            errors=errors,
        )

    async def async_step_delete_row(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        if not self._rows:
            self._warning = "Brak wierszy."
            return await self.async_step_rows_visual()
        if user_input is not None:
            drop = {int(i) for i in user_input.get("rows", [])}
            self._rows = [r for i, r in enumerate(self._rows) if i not in drop]
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
            item = self._rows.pop(int(user_input["row"]))
            pos = max(1, min(int(user_input["position"]), len(self._rows) + 1))
            self._rows.insert(pos - 1, item)
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

    async def async_step_rows_save(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        columns = list(self._opts.get(CONF_COLUMNS, DEFAULT_COLUMNS))
        err = layout_error(columns, self._rows)
        if err and ERR_OPIS in err:
            # wiersze z samym opisem wymagają kolumny "opis" - dodajemy ją automatycznie
            columns.insert(1 if columns and columns[0] == "lp" else 0, "opis")
            self._opts[CONF_COLUMNS] = columns
            err = layout_error(columns, self._rows)
        if err:
            self._warning = f"Nie można zapisać: {err}"
            return await self.async_step_rows_visual()
        self._opts[CONF_ROWS] = self._rows
        return self.async_create_entry(title="", data=self._opts)

    # ---- edycja tekstowa (YAML)
    async def async_step_layout_text(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        details = ""
        text = yaml.safe_dump(
            {
                CONF_COLUMNS: self._opts.get(CONF_COLUMNS, DEFAULT_COLUMNS),
                CONF_ROWS: self._rows,
            },
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
                    self._opts[CONF_COLUMNS] = _plain(valid[CONF_COLUMNS])
                    self._opts[CONF_ROWS] = _plain(valid[CONF_ROWS])
                    return self.async_create_entry(title="", data=self._opts)
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
            description_placeholders={"details": details},
        )
