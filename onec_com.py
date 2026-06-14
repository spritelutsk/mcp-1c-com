"""Подключение к 1С через COM (V83.COMConnector) и выполнение запросов.

Весь COM-доступ изолирован в одном выделенном потоке (single-thread apartment),
потому что COM-объекты 1С привязаны к создавшему их апартаменту и не переносятся
между потоками. Любая операция отправляется в этот поток через executor.
"""

from __future__ import annotations

import datetime as _dt
import threading
from concurrent.futures import ThreadPoolExecutor
from typing import Any

import pythoncom
import pywintypes
import win32com.client

from config import Settings


class OneCError(RuntimeError):
    """Ошибка работы с 1С через COM, пригодная для показа пользователю."""


def _friendly_com_error(exc: Exception, progid: str) -> OneCError:
    """Превращает невнятную COM-ошибку в подсказку с вероятной причиной."""
    text = str(exc)
    hints = []
    low = text.lower()
    if "не зарегистрирован" in low or "class not registered" in low or "80040154" in low:
        hints.append(
            f"COM-объект '{progid}' не зарегистрирован. Зарегистрируйте comcntr.dll нужной "
            f"разрядности (от администратора), например:\n"
            f'    regsvr32 "C:\\Program Files\\1cv8\\<версия>\\bin\\comcntr.dll"'
        )
    if "класс не зарегистрирован" in low:
        hints.append(
            "Разрядность Python и 1С должны совпадать: 64-битный Python ↔ 64-битный comcntr.dll, "
            "32-битный ↔ 32-битный."
        )
    msg = f"Ошибка COM-подключения к 1С: {text}"
    if hints:
        msg += "\n\nВозможная причина:\n" + "\n".join(hints)
    return OneCError(msg)


class OneCConnection:
    """Постоянное COM-подключение к информационной базе 1С.

    Подключение ленивое: устанавливается при первом обращении и переиспользуется.
    """

    def __init__(self, settings: Settings):
        self._settings = settings
        self._connection = None  # внешнее соединение (глобальный контекст 1С)
        self._executor_obj = None  # загруженная обработка-исполнитель кода
        self._lock = threading.Lock()
        # Один-единственный COM-апартамент на всё время жизни сервера.
        self._executor = ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="onec-com",
            initializer=self._init_thread,
        )

    @staticmethod
    def _init_thread() -> None:
        pythoncom.CoInitialize()

    def _run(self, fn, *args):
        """Выполнить функцию в COM-потоке и дождаться результата."""
        return self._executor.submit(fn, *args).result()

    # ── установка соединения ────────────────────────────────────────────────

    def _ensure_connection(self):
        if self._connection is not None:
            return self._connection
        progid = self._settings.com_connector
        conn_str = self._settings.connection_string()
        try:
            connector = win32com.client.Dispatch(progid)
        except Exception as exc:  # noqa: BLE001
            raise _friendly_com_error(exc, progid) from exc
        try:
            self._connection = connector.Connect(conn_str)
        except Exception as exc:  # noqa: BLE001
            safe = self._settings.connection_string(mask_password=True)
            raise OneCError(
                f"Не удалось подключиться к базе. Строка соединения: {safe}\n{exc}"
            ) from exc
        return self._connection

    def connect(self) -> None:
        """Принудительно установить соединение (для проверки конфигурации)."""
        self._run(self._ensure_connection)

    def close(self) -> None:
        def _close():
            self._executor_obj = None
            self._connection = None  # COM освободит соединение при сборке мусора

        self._run(_close)
        self._executor.shutdown(wait=True)

    # ── конвертация значений 1С → Python/JSON ──────────────────────────────

    def _to_python(self, value: Any) -> Any:
        if value is None:
            return None
        if isinstance(value, bool):  # bool раньше int!
            return value
        if isinstance(value, (int, float, str)):
            return value
        if isinstance(value, (_dt.datetime, pywintypes.TimeType)):
            try:
                dt = _dt.datetime(
                    value.year, value.month, value.day,
                    value.hour, value.minute, value.second,
                )
            except Exception:  # noqa: BLE001
                return str(value)
            # «Пустая дата» 1С (0001-01-01) — отдаём как null.
            if dt.year == 1 and dt.month == 1 and dt.day == 1:
                return None
            return dt.isoformat(sep=" ")
        # Прочее — это COM-объект (ссылка, перечисление, составной тип и т.п.).
        conn = self._connection
        try:
            # XMLСтрока вернёт UUID для ссылок и текстовое значение для примитивов.
            return conn.XMLСтрока(value)
        except Exception:  # noqa: BLE001
            try:
                return str(value)
            except Exception:  # noqa: BLE001
                return "<значение 1С>"

    def _py_to_param(self, value: Any) -> Any:
        """Готовит значение параметра Python → 1С."""
        if isinstance(value, str):
            # Похоже на ISO-дату? Передадим как дату 1С.
            dt = _try_parse_iso_date(value)
            if dt is not None:
                return pywintypes.Time(dt)
        return value

    # ── операции ────────────────────────────────────────────────────────────

    def execute_query(
        self,
        text: str,
        params: dict[str, Any] | None = None,
        limit: int = 1000,
    ) -> dict[str, Any]:
        return self._run(self._do_query, text, params or {}, limit)

    def _do_query(self, text: str, params: dict[str, Any], limit: int) -> dict[str, Any]:
        conn = self._ensure_connection()
        query = conn.NewObject("Запрос")
        query.Текст = text
        for name, raw in params.items():
            query.УстановитьПараметр(name, self._py_to_param(raw))
        try:
            result = query.Выполнить()
        except Exception as exc:  # noqa: BLE001
            raise OneCError(f"Ошибка выполнения запроса:\n{exc}") from exc

        table = result.Выгрузить()
        columns_coll = table.Колонки
        col_count = columns_coll.Количество()
        columns = [columns_coll.Получить(i).Имя for i in range(col_count)]

        total = table.Количество()
        take = total if limit is None or limit < 0 else min(total, limit)
        rows: list[list[Any]] = []
        for i in range(take):
            row = table.Получить(i)
            rows.append([self._to_python(row.Получить(j)) for j in range(col_count)])

        return {
            "columns": columns,
            "rows": rows,
            "row_count": len(rows),
            "total_rows": total,
            "truncated": take < total,
        }

    def list_metadata(self, kind: str | None = None) -> dict[str, Any]:
        return self._run(self._do_list_metadata, kind)

    # Сопоставление англ. ключей с коллекциями метаданных глобального контекста.
    _META_KINDS = {
        "catalogs": "Справочники",
        "documents": "Документы",
        "informationregisters": "РегистрыСведений",
        "accumulationregisters": "РегистрыНакопления",
        "enums": "Перечисления",
        "constants": "Константы",
        "chartsofcharacteristictypes": "ПланыВидовХарактеристик",
        "businessprocesses": "БизнесПроцессы",
        "tasks": "Задачи",
        "exchangeplans": "ПланыОбмена",
        "reports": "Отчеты",
        "dataprocessors": "Обработки",
    }

    def _do_list_metadata(self, kind: str | None) -> dict[str, Any]:
        conn = self._ensure_connection()
        meta = conn.Метаданные

        if kind:
            key = kind.strip().lower().replace(" ", "")
            ru = self._META_KINDS.get(key)
            if ru is None:
                raise OneCError(
                    f"Неизвестный вид метаданных '{kind}'. Доступно: "
                    + ", ".join(sorted(self._META_KINDS))
                )
            kinds = {key: ru}
        else:
            kinds = self._META_KINDS

        out: dict[str, list[str]] = {}
        for eng_key, ru_attr in kinds.items():
            collection = getattr(meta, ru_attr)
            names = []
            count = collection.Количество()
            for i in range(count):
                obj = collection.Получить(i)
                names.append(obj.Имя)
            out[eng_key] = names
        return {"metadata": out}

    def info(self) -> dict[str, Any]:
        return self._run(self._do_info)

    def _do_info(self) -> dict[str, Any]:
        conn = self._ensure_connection()
        md = conn.Метаданные
        config = md  # корень = метаданные конфигурации
        sysinfo = conn.NewObject("СистемнаяИнформация")
        return {
            "connected": True,
            "configuration_name": config.Имя,
            "configuration_synonym": str(config.Синоним),
            "configuration_version": config.Версия,
            "platform_version": sysinfo.ВерсияПриложения,
            "connection": self._settings.connection_string(mask_password=True),
        }

    # ── описание объекта метаданных ────────────────────────────────────────

    def describe_object(self, full_name: str) -> dict[str, Any]:
        return self._run(self._do_describe, full_name)

    def _format_type(self, td: Any) -> str:
        """Читаемое представление ОписанияТипов реквизита.

        Глобальный конструктор Тип() в старых режимах совместимости через COM
        недоступен, поэтому: ссылочные типы определяем через Метаданные.НайтиПоТипу
        (точно), а примитивы — по квалификаторам (Строка(N), Число(d,f)).
        ВАЖНО: ПолноеИмя() в 1С — метод (со скобками), а не свойство.
        """
        conn = self._connection
        ref_parts: list[str] = []
        total = 0
        try:
            types_arr = td.Типы()
            total = types_arr.Количество()
            for i in range(total):
                t = types_arr.Получить(i)
                try:
                    md = conn.Метаданные.НайтиПоТипу(t)
                except Exception:  # noqa: BLE001
                    md = None
                if md is not None:
                    ref_parts.append(md.ПолноеИмя())
        except Exception:  # noqa: BLE001
            pass

        prim_parts: list[str] = []
        try:
            length = td.КвалификаторыСтроки.Длина
            if length:
                prim_parts.append(f"Строка({length})")
        except Exception:  # noqa: BLE001
            pass
        try:
            digits = td.КвалификаторыЧисла.Разрядность
            if digits:
                frac = 0
                try:
                    frac = td.КвалификаторыЧисла.РазрядностьДробнойЧасти
                except Exception:  # noqa: BLE001
                    pass
                prim_parts.append(f"Число({digits},{frac})" if frac else f"Число({digits})")
        except Exception:  # noqa: BLE001
            pass

        parts = ref_parts + prim_parts
        # Остались неопознанные примитивы (Дата, Булево, неограниченная Строка,
        # ХранилищеЗначения) — их имена через COM в этом режиме не получить.
        if total and len(parts) < total:
            parts.append("примитив")
        return ", ".join(p for p in parts if p) if parts else "—"

    @staticmethod
    def _to_list(coll: Any) -> list[Any]:
        """Превращает COM-коллекцию метаданных в список.

        Обычные коллекции поддерживают Количество()/Получить(i), но особые
        (например СтандартныеРеквизиты) — только перебор через COM-итератор.
        """
        if coll is None:
            return []
        try:
            count = coll.Количество()
            return [coll.Получить(i) for i in range(count)]
        except Exception:  # noqa: BLE001
            pass
        try:
            return list(coll)
        except Exception:  # noqa: BLE001
            return []

    def _attrs(self, coll: Any) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for a in self._to_list(coll):
            item: dict[str, Any] = {"name": a.Имя}
            try:
                syn = str(a.Синоним)
                if syn:
                    item["synonym"] = syn
            except Exception:  # noqa: BLE001
                pass
            try:
                item["type"] = self._format_type(a.Тип)
            except Exception:  # noqa: BLE001
                pass
            out.append(item)
        return out

    def _do_describe(self, full_name: str) -> dict[str, Any]:
        conn = self._ensure_connection()
        try:
            md = conn.Метаданные.НайтиПоПолномуИмени(full_name)
        except Exception as exc:  # noqa: BLE001
            raise OneCError(f"Ошибка поиска объекта '{full_name}':\n{exc}") from exc
        if md is None:
            raise OneCError(
                f"Объект метаданных не найден: '{full_name}'. Укажите полное имя, "
                "например 'Справочник.Контрагенты' или 'Документ.РеализацияТоваровУслуг'."
            )

        res: dict[str, Any] = {
            "full_name": full_name,
            "name": md.Имя,
            "kind": full_name.split(".")[0],
        }
        try:
            res["synonym"] = str(md.Синоним)
        except Exception:  # noqa: BLE001
            pass

        std = self._attrs(getattr(md, "СтандартныеРеквизиты", None))
        attrs = self._attrs(getattr(md, "Реквизиты", None))
        dims = self._attrs(getattr(md, "Измерения", None))
        resources = self._attrs(getattr(md, "Ресурсы", None))

        tabs: list[dict[str, Any]] = []
        for t in self._to_list(getattr(md, "ТабличныеЧасти", None)):
            entry: dict[str, Any] = {"name": t.Имя}
            try:
                entry["synonym"] = str(t.Синоним)
            except Exception:  # noqa: BLE001
                pass
            entry["attributes"] = self._attrs(getattr(t, "Реквизиты", None))
            tabs.append(entry)

        if std:
            res["standard_attributes"] = std
        if attrs:
            res["attributes"] = attrs
        if dims:
            res["dimensions"] = dims
        if resources:
            res["resources"] = resources
        if tabs:
            res["tabular_sections"] = tabs
        return res

    # ── выполнение произвольного кода 1С ───────────────────────────────────
    #
    # Внешнее COM-соединение НЕ отдаёт операторы Выполнить/Вычислить (это
    # конструкции языка, а не методы глобального контекста). Поэтому код
    # исполняется внутри модуля внешней обработки-исполнителя (mcpИсполнитель.epf),
    # где Выполнить/Вычислить доступны как обычные операторы. Это же решает
    # возврат значения: код пишет в локальную переменную Результат функции.

    def _get_executor(self):
        if self._executor_obj is not None:
            return self._executor_obj
        path = self._settings.executor_epf
        if not path:
            raise OneCError(
                "Не задан путь к обработке-исполнителю. Укажите ONEC_EXECUTOR_EPF "
                "(путь к mcpИсполнитель.epf). По умолчанию ищется рядом с сервером."
            )
        import os

        if not os.path.exists(path):
            raise OneCError(
                f"Файл обработки-исполнителя не найден: {path}. Соберите его "
                "(executor/mcpИсполнитель.epf) или укажите ONEC_EXECUTOR_EPF."
            )
        conn = self._ensure_connection()
        try:
            self._executor_obj = conn.ВнешниеОбработки.Создать(path)
        except Exception as exc:  # noqa: BLE001
            text = str(exc)
            if "Предупреждение безопасности" in text or "Разрешить открывать" in text:
                raise OneCError(
                    "Платформа заблокировала загрузку обработки-исполнителя "
                    "«Предупреждением безопасности» (защита от опасных действий).\n"
                    "Чтобы execute_script работал, для пользователя ИБ нужно отключить "
                    "«Защиту от опасных действий»:\n"
                    "  Конфигуратор → Администрирование → Пользователи → <пользователь> "
                    "→ снять флаг «Защита от опасных действий».\n"
                    "Либо один раз откройте файл\n"
                    f"  {path}\n"
                    "в толстом клиенте 1С и подтвердите с галкой «Запомнить выбор»."
                ) from exc
            raise OneCError(
                f"Не удалось загрузить обработку-исполнитель '{path}':\n{exc}"
            ) from exc
        return self._executor_obj

    def execute_script(
        self, code: str | None, return_expression: str | None = None
    ) -> dict[str, Any]:
        return self._run(self._do_script, code, return_expression)

    def _do_script(
        self, code: str | None, return_expression: str | None
    ) -> dict[str, Any]:
        if not (code and code.strip()) and not (
            return_expression and return_expression.strip()
        ):
            raise OneCError("Не задан ни код (code), ни выражение (return_expression).")
        executor = self._get_executor()
        out: dict[str, Any] = {}
        if code and code.strip():
            try:
                value = executor.ВыполнитьКод(code)
            except Exception as exc:  # noqa: BLE001
                raise OneCError(f"Ошибка выполнения кода:\n{exc}") from exc
            out["executed"] = True
            # Результат, записанный кодом в переменную Результат, тоже отдаём.
            converted = self._to_python(value)
            if converted is not None:
                out["result"] = converted
        if return_expression and return_expression.strip():
            try:
                value = executor.ВычислитьВыражение(return_expression)
            except Exception as exc:  # noqa: BLE001
                raise OneCError(f"Ошибка вычисления выражения:\n{exc}") from exc
            out["result"] = self._to_python(value)
        return out


def _try_parse_iso_date(value: str) -> _dt.datetime | None:
    value = value.strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return _dt.datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None
