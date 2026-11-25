"""HA_mailer jest konfigurowany wyłącznie w configuration.yaml.

Ten plik jest potrzebny tylko do tego, by wpisy tworzone automatycznie z YAML (import)
były widoczne w Home Assistancie. Nie ma tu formularzy ani opcji.
"""
from __future__ import annotations

import json
from typing import Any

from homeassistant.config_entries import ConfigFlow, ConfigFlowResult

from . import CONF_ID, CONF_YAML_HASH, DOMAIN, yaml_digest


class HAMailerConfigFlow(ConfigFlow, domain=DOMAIN):
    VERSION = 3

    async def async_step_user(self, user_input: dict[str, Any] | None = None) -> ConfigFlowResult:
        """Przycisk „Dodaj integrację”: tylko informacja, że konfiguracja jest w YAML."""
        return self.async_abort(reason="yaml_only")

    async def async_step_import(self, import_data: dict[str, Any]) -> ConfigFlowResult:
        """Tworzy wpis dla bloku z YAML (wywoływane automatycznie przy starcie i po reload)."""
        conf = json.loads(json.dumps(import_data, default=str))
        await self.async_set_unique_id(f"yaml:{conf[CONF_ID]}")
        self._abort_if_unique_id_configured()
        return self.async_create_entry(
            title=f"HA_mailer ({conf[CONF_ID]})",
            data={**conf, CONF_YAML_HASH: yaml_digest(conf)},
        )
