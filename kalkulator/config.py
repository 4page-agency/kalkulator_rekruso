"""Zarządzanie konfiguracją i ustawieniami aplikacji."""

from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import json
import os
import sys
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict

CONFIG_DIR_NAME = "kalkulator_retruso"
CONFIG_FILE_NAME = "config.json"
DEFAULT_MARGIN_RULES = [
    {"max_quantity": 300, "margin_percent": 100.0},
    {"max_quantity": 500, "margin_percent": 70.0},
    {"max_quantity": 1000, "margin_percent": 25.0},
]
PBKDF2_ITERATIONS = 120_000


def _get_config_dir() -> Path:
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / CONFIG_DIR_NAME
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / CONFIG_DIR_NAME
    return Path.home() / ".config" / CONFIG_DIR_NAME


class ConfigManager:
    """Odpowiada za wczytywanie i zapisywanie ustawień programu."""

    def __init__(self) -> None:
        self.config_dir = _get_config_dir()
        self.config_file = self.config_dir / CONFIG_FILE_NAME
        self.data: Dict[str, Any] = {
            "password": None,
            "margin_rules": deepcopy(DEFAULT_MARGIN_RULES),
        }
        self.load()

    # ------------------------------------------------------------------
    # Operacje na pliku konfiguracyjnym
    # ------------------------------------------------------------------
    def load(self) -> None:
        try:
            with self.config_file.open("r", encoding="utf-8") as file:
                raw_data = json.load(file)
        except FileNotFoundError:
            return
        except (json.JSONDecodeError, OSError):
            return

        password_value = raw_data.get("password")
        if isinstance(password_value, str):
            self.data["password"] = password_value
        elif isinstance(password_value, dict):
            salt = password_value.get("salt")
            hash_value = password_value.get("hash")
            iterations = password_value.get("iterations", PBKDF2_ITERATIONS)
            if salt and hash_value:
                try:
                    iterations_int = int(iterations)
                except (TypeError, ValueError):
                    iterations_int = PBKDF2_ITERATIONS
                self.data["password"] = {
                    "salt": str(salt),
                    "hash": str(hash_value),
                    "iterations": iterations_int,
                }

        margin_rules = self._sanitize_margin_rules(raw_data.get("margin_rules"))
        if margin_rules is not None:
            self.data["margin_rules"] = margin_rules

    def save(self) -> None:
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            with self.config_file.open("w", encoding="utf-8") as file:
                json.dump(
                    {
                        "password": self.data.get("password"),
                        "margin_rules": self.data.get("margin_rules", []),
                    },
                    file,
                    ensure_ascii=False,
                    indent=2,
                )
        except OSError:
            pass

    # ------------------------------------------------------------------
    # Obsługa hasła
    # ------------------------------------------------------------------
    def has_password(self) -> bool:
        password_info = self.data.get("password")
        if isinstance(password_info, str):
            return bool(password_info)
        if isinstance(password_info, dict):
            return bool(password_info.get("salt") and password_info.get("hash"))
        return False

    def set_password(self, password: str) -> None:
        self.data["password"] = password
        self.save()

    def verify_password(self, password: str) -> bool:
        password_info = self.data.get("password")
        if isinstance(password_info, str):
            return password == password_info
        if isinstance(password_info, dict):
            return self._verify_hashed_password(password_info, password)
        return False

    @staticmethod
    def _verify_hashed_password(password_info: Dict[str, Any], password: str) -> bool:
        salt_b64 = password_info.get("salt")
        hash_b64 = password_info.get("hash")
        iterations = password_info.get("iterations", PBKDF2_ITERATIONS)
        if not salt_b64 or not hash_b64:
            return False
        try:
            salt = base64.b64decode(str(salt_b64))
            stored_hash = base64.b64decode(str(hash_b64))
            iterations_int = int(iterations)
        except (binascii.Error, TypeError, ValueError):
            return False
        new_hash = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), salt, iterations_int
        )
        return hmac.compare_digest(new_hash, stored_hash)

    # ------------------------------------------------------------------
    # Obsługa konfiguracji marży
    # ------------------------------------------------------------------
    def get_margin_rules(self) -> list[dict[str, float]]:
        return deepcopy(self.data.get("margin_rules", []))

    def update_margin_rules(
        self, rules: list[dict[str, float]]
    ) -> list[dict[str, float]]:
        sanitized = self._sanitize_margin_rules(rules)
        if sanitized is None:
            sanitized = deepcopy(DEFAULT_MARGIN_RULES)
        self.data["margin_rules"] = sanitized
        self.save()
        return deepcopy(self.data["margin_rules"])

    # ------------------------------------------------------------------
    # Funkcje pomocnicze
    # ------------------------------------------------------------------
    def _sanitize_margin_rules(
        self, value: Any
    ) -> list[dict[str, float]] | None:
        if not isinstance(value, list):
            return None
        sanitized: list[dict[str, float]] = []
        for item in value:
            if not isinstance(item, dict):
                continue
            try:
                max_quantity = int(item.get("max_quantity"))
                margin_percent = float(item.get("margin_percent"))
            except (TypeError, ValueError):
                continue
            sanitized.append(
                {"max_quantity": max_quantity, "margin_percent": margin_percent}
            )
        sanitized.sort(key=lambda rule: rule["max_quantity"])
        return sanitized


__all__ = [
    "ConfigManager",
    "DEFAULT_MARGIN_RULES",
]
