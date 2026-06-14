"""Конфигурация подключения к 1С через переменные окружения.

Поддерживаются два способа задать соединение:

1. Готовая строка соединения целиком:
       ONEC_CONNECTION_STRING=File="C:\\bases\\acc";Usr="Администратор";Pwd="..."

2. По частям (тогда строка собирается автоматически):
   Файловая база:
       ONEC_FILE=C:\\bases\\acc
   Серверная база:
       ONEC_SRVR=server-1c           (можно server-1c:1541)
       ONEC_REF=accounting
   Общие (для обоих вариантов):
       ONEC_USR=Администратор
       ONEC_PWD=пароль

Дополнительно:
   ONEC_COMCONNECTOR=V83.COMConnector   (ProgID коннектора, по умолчанию V83)
"""

from __future__ import annotations

import os
from dataclasses import dataclass

# Путь к обработке-исполнителю по умолчанию — рядом с этим модулем.
_DEFAULT_EXECUTOR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "executor", "mcpИсполнитель.epf"
)


def _q(value: str) -> str:
    """Экранирует значение для строки соединения 1С (двойные кавычки удваиваются)."""
    return value.replace('"', '""')


@dataclass
class Settings:
    com_connector: str
    raw_connection_string: str | None
    file_path: str | None
    server: str | None
    ref: str | None
    user: str | None
    password: str | None
    executor_epf: str | None

    @classmethod
    def from_env(cls) -> "Settings":
        return cls(
            com_connector=os.environ.get("ONEC_COMCONNECTOR", "V83.COMConnector"),
            raw_connection_string=os.environ.get("ONEC_CONNECTION_STRING"),
            file_path=os.environ.get("ONEC_FILE"),
            server=os.environ.get("ONEC_SRVR"),
            ref=os.environ.get("ONEC_REF"),
            user=os.environ.get("ONEC_USR"),
            password=os.environ.get("ONEC_PWD"),
            executor_epf=os.environ.get("ONEC_EXECUTOR_EPF", _DEFAULT_EXECUTOR),
        )

    def connection_string(self, mask_password: bool = False) -> str:
        """Собирает строку соединения для COMConnector.Connect()."""
        if self.raw_connection_string:
            s = self.raw_connection_string
            if mask_password:
                # Грубая маскировка Pwd=... в готовой строке.
                import re

                s = re.sub(r'(Pwd\s*=\s*)"[^"]*"', r'\1"***"', s, flags=re.IGNORECASE)
            return s

        parts: list[str] = []
        if self.file_path:
            parts.append(f'File="{_q(self.file_path)}"')
        elif self.server and self.ref:
            parts.append(f'Srvr="{_q(self.server)}"')
            parts.append(f'Ref="{_q(self.ref)}"')
        else:
            raise ValueError(
                "Не задано соединение. Укажите ONEC_CONNECTION_STRING, либо ONEC_FILE "
                "(файловая база), либо ONEC_SRVR + ONEC_REF (серверная база)."
            )

        if self.user:
            parts.append(f'Usr="{_q(self.user)}"')
        if self.password:
            pwd = "***" if mask_password else _q(self.password)
            parts.append(f'Pwd="{pwd}"')

        return ";".join(parts)
