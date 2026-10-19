"""
MCP-сервер над данными ежемесячного отчёта.

Даёт модели доступ к уже рассчитанным данным — только на чтение:
  - архивные отчёты  output/архив/отчёт_ГГГГ-ММ.xlsx  (их сохраняет notebook)
  - архив рядов      data/*.csv

Сервер ничего не загружает из интернета и ничего не меняет в файлах.
Запускается не вручную, а скриптом ai_commentary.py.

Важно: в этом файле нельзя использовать print — по stdout сервер общается с клиентом.
"""

from pathlib import Path
from typing import Annotated, Literal

import pandas as pd
from pydantic import Field
from mcp.server import MCPServer


PROJECT_DIR = Path(__file__).resolve().parent
DATA_DIR = PROJECT_DIR / "data"
REPORTS_DIR = PROJECT_DIR / "output" / "архив"

server = MCPServer(
    name="market-report",
    instructions="Данные ежемесячного обзора рынка: макроэкономика и ДКП, денежный рынок, облигации. "
                 "Все числа уже рассчитаны; используй их как есть.",
    log_level="WARNING",
)

Month = Annotated[str, Field(description="Отчётный месяц в формате ГГГГ-ММ, например 2026-09")]

SECTIONS = Literal["Макроэкономика и ДКП", "Денежный рынок", "Облигационный рынок", "все"]

# Ряды для get_series: название -> (файл в data/, колонка, описание с единицами)
SERIES = {
    "key_rate":   ("key_rate",  "key_rate",   "Ключевая ставка, %"),
    "inflation":  ("inflation", "inflation",  "Инфляция г/г, %"),
    "usdrub":     ("usdrub",    "usdrub",     "Курс USD/RUB, руб."),
    "ruonia":     ("ruonia",    "ruonia",     "RUONIA, %"),
    "rusfar":     ("rusfar",    "rusfar",     "RUSFAR ON, %"),
    "liquidity":  ("liquidity", "liquidity",  "Ликвидность банковского сектора, млрд руб.: значение больше нуля — ДЕФИЦИТ, меньше нуля — ПРОФИЦИТ"),
    "roisfix_1y": ("roisfix",   "roisfix_1y", "ROISfix 1Y, %"),
    "yield_aaa":  ("corp_aaa",  "yield_aaa",  "Доходность индекса корп. облигаций AAA, %"),
    "yield_a":    ("corp_a",    "yield_a",    "Доходность индекса корп. облигаций A, %"),
}


# ---------- Вспомогательные функции ----------

def read_block(month, sheet, columns):
    """Блок из архивного отчёта: заголовок таблицы в строке 3, columns — диапазон колонок, например "A:J" """
    path = REPORTS_DIR / f"отчёт_{month}.xlsx"
    return pd.read_excel(path, sheet_name=sheet, header=2, usecols=columns).dropna(how="all")


def report_missing(month):
    """Если отчёта за месяц нет — понятное сообщение для модели, иначе None"""
    if (REPORTS_DIR / f"отчёт_{month}.xlsx").exists():
        return None
    return {"error": f"Нет отчёта за {month}", "available_months": available_months()}


def to_records(table):
    """Таблица -> список словарей: числа округлены до 2 знаков, даты — ДД.ММ.ГГГГ, пропуски — None"""
    table = table.copy()
    for column in table.columns:
        if pd.api.types.is_datetime64_any_dtype(table[column]):
            table[column] = table[column].dt.strftime("%d.%m.%Y")
        elif pd.api.types.is_numeric_dtype(table[column]):
            table[column] = table[column].round(2)
    table = table.astype(object).where(table.notna(), None)
    return table.to_dict(orient="records")


def describe_liquidity(value, previous):
    """Ликвидность словами, чтобы модели не нужно было расшифровывать знак.
    В данных ЦБ плюс — дефицит ликвидности, минус — профицит."""
    state = "дефицит" if value > 0 else "профицит"
    text = f"{state.capitalize()} ликвидности: {abs(value):.1f} млрд руб."
    if previous is not None:
        previous_state = "дефицит" if previous > 0 else "профицит"
        if previous_state != state:
            text += f" Месяц назад был {previous_state} {abs(previous):.1f} млрд руб."
        else:
            difference = abs(value) - abs(previous)
            direction = "увеличился" if difference > 0 else "сократился" if difference < 0 else "не изменился"
            text += f" За месяц {state} {direction}" + (f" на {abs(difference):.1f} млрд руб." if difference else ".")
    return state, text


def available_months():
    return sorted(path.stem.replace("отчёт_", "") for path in REPORTS_DIR.glob("отчёт_*.xlsx"))


# ---------- Инструменты ----------

@server.tool()
def list_reports() -> dict:
    """Список месяцев, за которые есть готовые отчёты."""
    return {"months": available_months()}


@server.tool()
def get_summary(
    month: Month,
    section: Annotated[SECTIONS, Field(description="Раздел отчёта или 'все'")] = "все",
) -> dict:
    """Сводная таблица отчёта: значение на отчётную дату, месяц назад, на начало года и изменения.
    Единицы изменений указаны в колонке 'Ед. изменения' (б.п., п.п., %).
    Колонка 'Актуально на' — дата последнего доступного значения (инфляция публикуется с лагом)."""
    if report_missing(month):
        return report_missing(month)
    summary = read_block(month, "Сводка", "A:J")
    if section != "все":
        summary = summary[summary["Раздел"] == section]
    rows = to_records(summary)

    # ликвидность — с расшифровкой знака
    for row in rows:
        if row["Показатель"].startswith("Ликвидность") and row["Значение"] is not None:
            state, explanation = describe_liquidity(row["Значение"], row["Месяц назад"])
            row["Состояние на отчётную дату"] = state
            row["Пояснение"] = explanation
    return {"month": month, "rows": rows}


@server.tool()
def get_kbd(month: Month) -> dict:
    """КБД ОФЗ на отчётную дату, месяц назад и начало года по срокам 0,25–20 лет; изменения в б.п."""
    if report_missing(month):
        return report_missing(month)
    curve = read_block(month, "Облигации", "A:F")
    curve_dates = read_block(month, "Облигации", "M:N")
    return {"month": month, "curve": to_records(curve), "curve_dates": to_records(curve_dates)}


@server.tool()
def get_credit_spreads(month: Month) -> dict:
    """Доходности индексов корпоративных облигаций AAA и A, доходность КБД на их дюрации
    и кредитные спреды к ОФЗ (б.п.) на концы последних 13 месяцев."""
    if report_missing(month):
        return report_missing(month)
    spreads = read_block(month, "Облигации", "Y:AE")
    return {"month": month, "rows": to_records(spreads)}


@server.tool()
def get_rate_expectations(month: Month) -> dict:
    """Ожидания по ключевой ставке: рыночная оценка средней ключевой ставки по ROISfix на сроки 1W–2Y
    (сейчас и месяц назад), а также последние прогнозы ЦБ и консенсус аналитиков (средняя ключевая за год)."""
    if report_missing(month):
        return report_missing(month)
    market = read_block(month, "Прогноз_ставки", "A:E")
    forecasts = read_block(month, "Прогноз_ставки", "M:Q")
    return {"month": month, "market_roisfix": to_records(market), "cbr_and_consensus": to_records(forecasts)}


@server.tool()
def get_series(
    indicator: Annotated[Literal[tuple(SERIES)], Field(description="Показатель: " + ", ".join(
        f"{name} — {description}" for name, (_, _, description) in SERIES.items()))],
    month: Month,
    months_back: Annotated[int, Field(description="Сколько последних месяцев показать (1–36)")] = 12,
) -> dict:
    """История показателя: значения на конец каждого месяца, заканчивая отчётным месяцем."""
    file_name, column, description = SERIES[indicator]
    table = pd.read_csv(DATA_DIR / f"{file_name}.csv", parse_dates=["date"]).dropna(subset=[column])

    month_ends = pd.date_range(end=pd.Period(month, freq="M").end_time.normalize(),
                               periods=max(1, min(months_back, 36)), freq="ME")
    values = [
        {"month_end": f"{date:%d.%m.%Y}",
         "value": round(float(table.loc[table["date"] <= date, column].iloc[-1]), 2)
                  if (table["date"] <= date).any() else None}
        for date in month_ends
    ]
    if indicator == "liquidity":            # знак ЦБ: плюс — дефицит, минус — профицит
        for item in values:
            if item["value"] is not None:
                item["state"] = "дефицит" if item["value"] > 0 else "профицит"
    return {"indicator": indicator, "description": description, "values": values}


@server.tool()
def get_data_quality(month: Month) -> dict:
    """Проверки качества данных и журнал загрузки: какие данные устарели или не загрузились."""
    if report_missing(month):
        return report_missing(month)
    checks = read_block(month, "Лог", "A:C")
    log = read_block(month, "Лог", "M:O")
    problems = log[log["Статус"] != "OK"]
    return {"month": month, "checks": to_records(checks), "problems": to_records(problems)}


if __name__ == "__main__":
    server.run("stdio")
