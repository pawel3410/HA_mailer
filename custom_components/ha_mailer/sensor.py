"""Encje HA_mailer: status, ostatnia wysyłka, statystyki."""
from __future__ import annotations

from homeassistant.components.sensor import SensorDeviceClass, SensorEntity, SensorStateClass
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.device_registry import DeviceInfo
from homeassistant.helpers.dispatcher import async_dispatcher_connect
from homeassistant.util import dt as dt_util

from . import CONF_RECIPIENT, CONF_ROWS, DOMAIN, ROW_ENTITY, SIGNAL, HAMailer


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry, async_add_entities) -> None:
    m: HAMailer = hass.data[DOMAIN][entry.entry_id]
    async_add_entities([StatusSensor(m), LastSentSensor(m), StatsSensor(m)])


class _Base(SensorEntity):
    _attr_has_entity_name = True
    _attr_should_poll = False
    _key = ""

    def __init__(self, mailer: HAMailer) -> None:
        self._m = mailer
        eid = mailer.entry.entry_id
        self._attr_unique_id = f"{eid}_{self._key}"
        self._attr_device_info = DeviceInfo(
            identifiers={(DOMAIN, eid)}, name="HA_mailer", manufacturer="HA_mailer"
        )

    async def async_added_to_hass(self) -> None:
        self.async_on_remove(
            async_dispatcher_connect(
                self.hass, SIGNAL.format(self._m.entry.entry_id), self.async_write_ha_state
            )
        )


class StatusSensor(_Base):
    _key = "status"
    _attr_name = "Status wysyłki"

    @property
    def native_value(self) -> str:
        return {"ok": "OK", "error": "Błąd"}.get(self._m.stats["status"], "Brak danych")

    @property
    def icon(self) -> str:
        return {"ok": "mdi:email-check", "error": "mdi:email-alert"}.get(
            self._m.stats["status"], "mdi:email-outline"
        )

    @property
    def extra_state_attributes(self) -> dict:
        s = self._m.stats
        return {
            "blad": s["last_error"] if s["status"] == "error" else None,
            "ostatnia_proba": s["last_attempt"],
            "odbiorca": self._m.cfg[CONF_RECIPIENT],
        }


class LastSentSensor(_Base):
    _key = "last_sent"
    _attr_name = "Ostatnia wysyłka"
    _attr_device_class = SensorDeviceClass.TIMESTAMP
    _attr_icon = "mdi:email-fast"

    @property
    def native_value(self):
        raw = self._m.stats["last_ok"]
        return dt_util.parse_datetime(raw) if raw else None

    @property
    def extra_state_attributes(self) -> dict:
        return {"ostatnia_proba": self._m.stats["last_attempt"]}


class StatsSensor(_Base):
    _key = "stats"
    _attr_name = "Statystyki"
    _attr_icon = "mdi:chart-box-outline"
    _attr_native_unit_of_measurement = "wysyłek"
    _attr_state_class = SensorStateClass.TOTAL_INCREASING

    @property
    def native_value(self) -> int:
        return self._m.stats["ok"]  # liczba udanych wysyłek

    @property
    def extra_state_attributes(self) -> dict:
        s = self._m.stats
        total = s["ok"] + s["fail"]
        rows = self._m.cfg[CONF_ROWS]
        return {
            "udane": s["ok"],
            "nieudane": s["fail"],
            "razem": total,
            "skutecznosc_proc": round(100 * s["ok"] / total, 1) if total else None,
            "ostatni_blad": s["last_error"],
            "czas_ostatniego_bledu": s["last_error_time"],
            "rozmiar_ostatniego_pliku_kb": s["last_file_kb"],
            "nastepna_wysylka": self._m.next_run.isoformat() if self._m.next_run else None,
            "liczba_encji": sum(1 for r in rows if r.get(ROW_ENTITY)),
            "liczba_wierszy": len(rows),
        }
