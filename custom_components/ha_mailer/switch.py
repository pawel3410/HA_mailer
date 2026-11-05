"""Przełącznik HA_mailer: włącza i wyłącza wysyłanie raportu z harmonogramu."""
from __future__ import annotations

from homeassistant.components.switch import SwitchEntity
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.dispatcher import async_dispatcher_connect

from . import DOMAIN, SIGNAL, HAMailer


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry, async_add_entities) -> None:
    m: HAMailer = hass.data[DOMAIN][entry.entry_id]
    async_add_entities([ScheduleSwitch(m, rid) for rid in m.reports])


class ScheduleSwitch(SwitchEntity):
    """Wyłączony: nie powstają nowe raporty z harmonogramu (zaległe pliki nadal są wysyłane)."""

    _attr_has_entity_name = True
    _attr_should_poll = False
    _attr_name = "Wysyłka z harmonogramu"

    def __init__(self, mailer: HAMailer, rid: str) -> None:
        self._m = mailer
        self._rid = rid
        self._attr_unique_id = f"{mailer.entry.entry_id}_{rid}_enabled"
        self._attr_device_info = mailer.device_info(rid)

    @property
    def is_on(self) -> bool:
        return self._m.enabled.get(self._rid, True)

    @property
    def icon(self) -> str:
        return "mdi:calendar-check" if self.is_on else "mdi:calendar-remove"

    @property
    def extra_state_attributes(self) -> dict:
        nxt = self._m.next_runs.get(self._rid)
        return {"nastepna_wysylka": nxt.isoformat() if nxt else None}

    async def async_turn_on(self, **kwargs) -> None:
        await self._m.async_set_enabled(self._rid, True)

    async def async_turn_off(self, **kwargs) -> None:
        await self._m.async_set_enabled(self._rid, False)

    async def async_added_to_hass(self) -> None:
        self.async_on_remove(
            async_dispatcher_connect(
                self.hass, SIGNAL.format(self._m.entry.entry_id), self.async_write_ha_state
            )
        )
