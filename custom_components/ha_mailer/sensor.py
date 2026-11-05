"""Encje HA_mailer: dla każdego raportu status, ostatnia wysyłka, statystyki + kolejka."""
from __future__ import annotations

from homeassistant.components.sensor import SensorDeviceClass, SensorEntity, SensorStateClass
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.dispatcher import async_dispatcher_connect
from homeassistant.util import dt as dt_util

from . import (
    CONF_NAME,
    CONF_RECIPIENT,
    CONF_RETENTION,
    CONF_ROWS,
    DEFAULT_RETENTION,
    DOMAIN,
    ROW_ENTITY,
    SIGNAL,
    HAMailer,
    count_sheets,
)


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry, async_add_entities) -> None:
    m: HAMailer = hass.data[DOMAIN][entry.entry_id]
    entities: list[SensorEntity] = [PendingSensor(m)]
    for rid in m.reports:
        entities += [StatusSensor(m, rid), LastSentSensor(m, rid), StatsSensor(m, rid)]
    async_add_entities(entities)


class _Base(SensorEntity):
    _attr_has_entity_name = True
    _attr_should_poll = False
    _key = ""

    def __init__(self, mailer: HAMailer, rid: str | None = None) -> None:
        self._m = mailer
        self._rid = rid
        eid = mailer.entry.entry_id
        self._attr_unique_id = f"{eid}_{self._key}" if rid is None else f"{eid}_{rid}_{self._key}"
        self._attr_device_info = mailer.device_info(rid)

    @property
    def _stats(self) -> dict:
        return self._m.stats[self._rid]

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
        return {"ok": "OK", "error": "Błąd"}.get(self._stats["status"], "Brak danych")

    @property
    def icon(self) -> str:
        return {"ok": "mdi:email-check", "error": "mdi:email-alert"}.get(
            self._stats["status"], "mdi:email-outline"
        )

    @property
    def extra_state_attributes(self) -> dict:
        s = self._stats
        return {
            "blad": s["last_error"] if s["status"] == "error" else None,
            "ostatnia_proba": s["last_attempt"],
            "oczekujace_pliki": len(self._m.pending_for(self._rid)),
            "aktywny": self._m.enabled.get(self._rid, True),
            "odbiorca": self._m.reports[self._rid][CONF_RECIPIENT],
        }


class LastSentSensor(_Base):
    _key = "last_sent"
    _attr_name = "Ostatnia wysyłka"
    _attr_device_class = SensorDeviceClass.TIMESTAMP
    _attr_icon = "mdi:email-fast"

    @property
    def native_value(self):
        raw = self._stats["last_ok"]
        return dt_util.parse_datetime(raw) if raw else None

    @property
    def extra_state_attributes(self) -> dict:
        return {"ostatnia_proba": self._stats["last_attempt"]}


class StatsSensor(_Base):
    _key = "stats"
    _attr_name = "Statystyki"
    _attr_icon = "mdi:chart-box-outline"
    _attr_native_unit_of_measurement = "wysyłek"
    _attr_state_class = SensorStateClass.TOTAL_INCREASING

    @property
    def native_value(self) -> int:
        return self._stats["ok"]  # liczba udanych wysyłek

    @property
    def extra_state_attributes(self) -> dict:
        s = self._stats
        total = s["ok"] + s["fail"]
        rows = self._m.reports[self._rid][CONF_ROWS]
        nxt = self._m.next_runs.get(self._rid)
        return {
            "udane": s["ok"],
            "udane_po_ponowieniu": s["retry_ok"],
            "nieudane_proby": s["fail"],
            "razem_prob": total,
            "skutecznosc_prob_proc": round(100 * s["ok"] / total, 1) if total else None,
            "ostatni_blad": s["last_error"],
            "czas_ostatniego_bledu": s["last_error_time"],
            "rozmiar_ostatniego_pliku_kb": s["last_file_kb"],
            "nastepna_wysylka": nxt.isoformat() if nxt else None,
            "liczba_encji": sum(1 for r in rows if r.get(ROW_ENTITY)),
            "liczba_wierszy": len(rows),
            "liczba_arkuszy": count_sheets(rows),
        }


class PendingSensor(_Base):
    _key = "pending"
    _attr_name = "Oczekujące wysyłki"
    _attr_icon = "mdi:email-sync"
    _attr_native_unit_of_measurement = "plików"

    @property
    def native_value(self) -> int:
        return len(self._m.pending)

    @property
    def extra_state_attributes(self) -> dict:
        m = self._m
        return {
            "kolejka": [
                {
                    "raport": m.reports.get(i["report"], {}).get(CONF_NAME, i["report"]),
                    "plik": i["file"],
                    "zapisany_na_dysku": i.get("stored", True),
                    "utworzono": i["created"],
                    "nieudane_proby": i["attempts"],
                    "nastepna_proba": i["next_try"],
                    "blad": i["last_error"],
                }
                for i in m.pending
            ],
            "katalog_plikow": m.folder,
            "katalog_dostepny": m.folder_ok,
            "pliki_w_katalogu": m.file_count,
            "retencja_dni": m.cfg.get(CONF_RETENTION, DEFAULT_RETENTION),
        }
