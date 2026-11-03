"""Config flow i opcje HA_mailer."""
from __future__ import annotations

import re
import smtplib
import ssl
from typing import Any

import voluptuous as vol
from croniter import croniter

from homeassistant.config_entries import ConfigEntry, ConfigFlow, ConfigFlowResult, OptionsFlow
from homeassistant.core import HomeAssistant, callback
from homeassistant.helpers import selector

from . import (
    CONF_CRON,
    CONF_ENTITIES,
    CONF_GMAIL_PASSWORD,
    CONF_GMAIL_USER,
    CONF_RECIPIENT,
    CONF_SUBJECT,
    DEFAULT_CRON,
    DEFAULT_SUBJECT,
    DOMAIN,
    GMAIL_HOST,
    GMAIL_PORT,
)

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


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


async def _validate(hass: HomeAssistant, d: dict[str, Any]) -> dict[str, str]:
    errors: dict[str, str] = {}
    if not d.get(CONF_ENTITIES):
        errors[CONF_ENTITIES] = "no_entities"
    if not croniter.is_valid(d.get(CONF_CRON, "")):
        errors[CONF_CRON] = "invalid_cron"
    if not EMAIL_RE.match(d.get(CONF_GMAIL_USER, "")):
        errors[CONF_GMAIL_USER] = "invalid_email"
    rcpt = [r.strip() for r in d.get(CONF_RECIPIENT, "").replace(";", ",").split(",") if r.strip()]
    if not rcpt or not all(EMAIL_RE.match(r) for r in rcpt):
        errors[CONF_RECIPIENT] = "invalid_recipient"
    if not d.get(CONF_SUBJECT):
        errors[CONF_SUBJECT] = "no_subject"
    if errors:
        return errors
    try:
        await hass.async_add_executor_job(_try_login, d[CONF_GMAIL_USER], d[CONF_GMAIL_PASSWORD])
    except smtplib.SMTPAuthenticationError:
        errors["base"] = "auth_failed"
    except (OSError, smtplib.SMTPException):
        errors["base"] = "cannot_connect"
    return errors


def _schema(defaults: dict[str, Any], password_required: bool) -> vol.Schema:
    pw = vol.Required(CONF_GMAIL_PASSWORD) if password_required else vol.Optional(CONF_GMAIL_PASSWORD)
    return vol.Schema(
        {
            vol.Required(CONF_ENTITIES, default=defaults.get(CONF_ENTITIES, [])): selector.EntitySelector(
                selector.EntitySelectorConfig(multiple=True)
            ),
            vol.Required(CONF_CRON, default=defaults.get(CONF_CRON, DEFAULT_CRON)): selector.TextSelector(),
            vol.Required(CONF_GMAIL_USER, default=defaults.get(CONF_GMAIL_USER, "")): selector.TextSelector(
                selector.TextSelectorConfig(type=selector.TextSelectorType.EMAIL)
            ),
            pw: selector.TextSelector(
                selector.TextSelectorConfig(type=selector.TextSelectorType.PASSWORD)
            ),
            vol.Required(CONF_RECIPIENT, default=defaults.get(CONF_RECIPIENT, "")): selector.TextSelector(),
            vol.Required(CONF_SUBJECT, default=defaults.get(CONF_SUBJECT, DEFAULT_SUBJECT)): selector.TextSelector(),
        }
    )


class HAMailerConfigFlow(ConfigFlow, domain=DOMAIN):
    VERSION = 1

    @staticmethod
    @callback
    def async_get_options_flow(config_entry: ConfigEntry) -> OptionsFlow:
        return HAMailerOptionsFlow()

    async def async_step_user(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        if user_input is not None:
            user_input = _normalize(user_input)
            errors = await _validate(self.hass, user_input)
            if not errors:
                return self.async_create_entry(title="HA_mailer", data=user_input)
        defaults = {k: v for k, v in (user_input or {}).items() if k != CONF_GMAIL_PASSWORD}
        return self.async_show_form(
            step_id="user", data_schema=_schema(defaults, True), errors=errors
        )


class HAMailerOptionsFlow(OptionsFlow):
    async def async_step_init(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        errors: dict[str, str] = {}
        current = {**self.config_entry.data, **self.config_entry.options}
        if user_input is not None:
            user_input = _normalize(user_input)
            if not user_input.get(CONF_GMAIL_PASSWORD):
                user_input[CONF_GMAIL_PASSWORD] = current[CONF_GMAIL_PASSWORD]
            # stała pozycja encji: istniejące zostają na swoich miejscach, nowe dopisujemy na końcu
            chosen = user_input.get(CONF_ENTITIES, [])
            old = current.get(CONF_ENTITIES, [])
            user_input[CONF_ENTITIES] = [e for e in old if e in chosen] + [
                e for e in chosen if e not in old
            ]
            errors = await _validate(self.hass, user_input)
            if not errors:
                return self.async_create_entry(title="", data=user_input)
        defaults = {k: v for k, v in (user_input or current).items() if k != CONF_GMAIL_PASSWORD}
        return self.async_show_form(
            step_id="init", data_schema=_schema(defaults, False), errors=errors
        )
