"""MCP-сервер для 1С через COM-соединение (V83.COMConnector).

Работает только на Windows с установленной платформой 1С:Предприятие.
Запуск (stdio-транспорт):

    python server.py

Конфигурация — через переменные окружения (см. config.py / README.md).
"""

from __future__ import annotations

import json
from typing import Any

from mcp.server.fastmcp import FastMCP

from config import Settings
from onec_com import OneCConnection, OneCError

settings = Settings.from_env()
conn = OneCConnection(settings)

mcp = FastMCP(
    "1c-com",
    instructions=(
        "Сервер выполняет запросы и читает метаданные информационной базы 1С через "
        "COM-соединение. Используй execute_query для запросов на встроенном языке "
        "запросов 1С (русские ключевые слова: ВЫБРАТЬ, ИЗ, ГДЕ, ...). "
        "Ссылочные поля возвращаются как UUID — чтобы получить читаемое представление, "
        "добавляй в запрос ПРЕДСТАВЛЕНИЕ(Поле) КАК ПолеПредставление."
    ),
)


def _err(exc: Exception) -> str:
    if isinstance(exc, (OneCError, ValueError)):
        return f"ОШИБКА: {exc}"
    return f"ОШИБКА ({type(exc).__name__}): {exc}"


@mcp.tool()
def execute_query(
    query_text: str,
    params: dict[str, Any] | None = None,
    limit: int = 1000,
) -> str:
    """Выполнить запрос на языке запросов 1С и вернуть результат.

    Args:
        query_text: Текст запроса (ВЫБРАТЬ ... ИЗ ... ГДЕ ...).
        params: Параметры запроса {ИмяПараметра: значение}. Строки вида
            "2026-06-12" или "2026-06-12T10:00:00" передаются как дата 1С.
        limit: Максимум строк в ответе (по умолчанию 1000; -1 — без ограничения).

    Returns:
        JSON: {columns, rows, row_count, total_rows, truncated}.
        Ссылочные значения отдаются как UUID — для читаемого вида используйте
        ПРЕДСТАВЛЕНИЕ(...) в самом запросе.
    """
    try:
        result = conn.execute_query(query_text, params, limit)
        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def list_metadata(kind: str | None = None) -> str:
    """Список объектов метаданных конфигурации.

    Args:
        kind: Вид метаданных (англ. ключ): catalogs, documents,
            informationregisters, accumulationregisters, enums, constants,
            reports, dataprocessors, exchangeplans, tasks, businessprocesses,
            chartsofcharacteristictypes. Без указания — все виды.

    Returns:
        JSON {metadata: {вид: [имена объектов]}}.
    """
    try:
        return json.dumps(conn.list_metadata(kind), ensure_ascii=False)
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def describe_object(full_name: str) -> str:
    """Описать объект метаданных: реквизиты, табличные части, типы.

    Args:
        full_name: Полное имя объекта, например "Справочник.Контрагенты",
            "Документ.РеализацияТоваровУслуг", "РегистрСведений.ЦеныНоменклатуры".

    Returns:
        JSON: full_name, name, kind, synonym, standard_attributes, attributes,
        tabular_sections (с реквизитами), а для регистров — dimensions/resources.
        Типы реквизитов даны в читаемом виде (ссылочные — полным именем объекта,
        примитивные — с квалификаторами, напр. "Строка(50)", "Число(15,2)").
    """
    try:
        return json.dumps(conn.describe_object(full_name), ensure_ascii=False)
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def execute_script(
    code: str | None = None,
    return_expression: str | None = None,
) -> str:
    """Выполнить произвольный код на встроенном языке 1С (серверный контекст).

    Код выполняется во внешнем соединении через глобальный метод `Выполнить`.
    Чтобы вернуть значение, задайте `return_expression` — оно вычисляется через
    `Вычислить`.

    ВАЖНО: `Выполнить` и `Вычислить` исполняются в РАЗНЫХ контекстах — переменные,
    созданные в `code`, не видны из `return_expression`. Поэтому возвращаемое
    значение вычисляйте самостоятельным выражением (обращение к справочникам,
    документам, регистрам), а не через переменную из code. Пример:
        return_expression = "Справочники.Валюты.НайтиПоКоду(\"980\").Наименование"

    Args:
        code: Операторы 1С для выполнения (побочные действия: запись объектов,
            проведение документов и т.п.). Необязательно.
        return_expression: Выражение 1С, результат которого вернуть. Необязательно.
            Должно быть задано хотя бы одно из code / return_expression.

    Returns:
        JSON {executed?: true, result?: <значение>}. Ссылки — как UUID.
    """
    try:
        return json.dumps(
            conn.execute_script(code, return_expression), ensure_ascii=False, default=str
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def connection_info() -> str:
    """Проверить соединение и получить сведения о базе и платформе.

    Returns:
        JSON с именем и версией конфигурации, версией платформы и (замаскированной)
        строкой соединения. При ошибке — текст с вероятной причиной.
    """
    try:
        return json.dumps(conn.info(), ensure_ascii=False)
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


if __name__ == "__main__":
    mcp.run()
