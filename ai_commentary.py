"""
ИИ-комментарии к ежемесячному отчёту (учебная часть проекта).

Что делает скрипт:
  1. Запускает MCP-сервер (mcp_server.py) с данными отчёта.
  2. Для каждого раздела отчёта просит GigaChat написать тезисы. Модель сама решает,
     какие данные ей нужны, и запрашивает их через инструменты MCP-сервера.
  3. Код проверяет тезисы: каждое число должно быть в данных, которые модель получила,
     а состояние ликвидности (дефицит/профицит) — совпадать с данными.
     Если нет — возвращает модели замечания, и она переписывает тезисы (до 2 попыток).
  4. Записывает результат в output/тезисы_ГГГГ-ММ.xlsx.

Запуск (обычно — из последнего раздела notebook):
  python ai_commentary.py 2026-09      — тезисы за сентябрь 2026
  python ai_commentary.py --check      — только проверить подключение к GigaChat

Перед первым запуском нужны:
  - файл gigachat_key.txt с ключом авторизации GigaChat API (рядом с этим скриптом);
  - корневой сертификат Минцифры в папке certs/ (используется только этим скриптом).
"""

import sys
import ssl
import json
import re
import asyncio
from pathlib import Path
from datetime import datetime

import certifi
import pandas as pd
from gigachat import GigaChat
from gigachat.models import Chat, Messages, MessagesRole, Function
from openpyxl.styles import Alignment
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


# =====================================================================
# Параметры
# =====================================================================
PROJECT_DIR = Path(__file__).resolve().parent

CHECK_ONLY = "--check" in sys.argv
REPORT_MONTH = next((arg for arg in sys.argv[1:] if not arg.startswith("--")), None)

MODEL = "GigaChat-2-Pro"            # можно GigaChat-2 (Lite) или GigaChat-2-Max
TEMPERATURE = 0.2                   # чем ниже, тем меньше «творчества»
MAX_TOOL_CALLS = 8                  # сколько раз модель может обратиться к данным в одном разделе
MAX_FIX_ATTEMPTS = 2                # сколько раз можно исправлять тезисы после проверки чисел

KEY_FILE = PROJECT_DIR / "gigachat_key.txt"
CERTS_DIR = PROJECT_DIR / "certs"
CA_BUNDLE = PROJECT_DIR / "output" / "_ca_bundle_gigachat.pem"
OUTPUT_FILE = PROJECT_DIR / "output" / f"тезисы_{REPORT_MONTH}.xlsx"

# Разделы отчёта и о чём в них писать
SECTIONS = {
    "Макроэкономика и ДКП": "ключевая ставка, инфляция и реальная ставка, курс доллара, "
                            "ожидания по ключевой ставке (оценка рынка по ROISfix, прогноз ЦБ, консенсус)",
    "Денежный рынок":       "RUONIA и RUSFAR, их спреды к ключевой ставке, ликвидность банковского сектора",
    "Облигационный рынок":  "КБД ОФЗ по срокам, доходности корпоративных облигаций AAA и A, "
                            "кредитные спреды к ОФЗ",
}

SYSTEM_PROMPT = """Ты — аналитик инвестиционного подразделения страховой компании. Ты пишешь тезисы
к ежемесячному обзору рынка для финансового директора, генерального директора и правления.

Правила:
1. Данные получай только через функции. Каждое число в тексте должно быть взято из ответа функции.
   Не вычисляй новые числа сам (разницы, суммы, средние) — нужные изменения уже есть в данных.
2. Числа можно округлять не больше чем до одного знака после запятой. Единицы — как в данных:
   изменения ставок и доходностей — в б.п., инфляции — в п.п., курса — в %, ликвидность — в млрд руб.
3. Если показатель актуален не на отчётную дату (например, инфляция публикуется с лагом), укажи, за какой он месяц.
4. Пиши 3–4 коротких тезиса, каждый с новой строки, начиная с «• ». Только факты и их смысл
   для рынка, без собственных прогнозов и инвестиционных рекомендаций.
5. Пиши по-русски, деловым языком, без вводных фраз.
6. Ликвидность банковского сектора: в данных ЦБ значение больше нуля — это ДЕФИЦИТ ликвидности,
   меньше нуля — ПРОФИЦИТ. Рост дефицита означает, что ликвидности стало меньше.
   Опирайся на готовые поля «Состояние на отчётную дату» и «Пояснение», а не расшифровывай знак сам."""


# =====================================================================
# Подключение к GigaChat
# =====================================================================
def make_ca_bundle():
    """Файл доверенных сертификатов только для клиента GigaChat:
    стандартный набор (certifi) + сертификаты Минцифры из папки certs/.
    В систему и в общий набор Python сертификаты не добавляются."""
    russian_certs = [path for path in sorted(CERTS_DIR.glob("*")) if path.suffix.lower() in {".crt", ".cer", ".pem"}]
    if not russian_certs:
        raise SystemExit(
            f"Нет сертификатов в папке {CERTS_DIR}.\n"
            "Скачай корневой сертификат Минцифры с https://www.gosuslugi.ru/crt и положи файл в папку certs/.")

    pem_texts = []
    for path in russian_certs:
        content = path.read_bytes()
        if b"BEGIN CERTIFICATE" in content:
            pem_texts.append(content.decode("ascii"))
        else:
            pem_texts.append(ssl.DER_cert_to_PEM_cert(content))      # файл в двоичном формате (DER)

    CA_BUNDLE.parent.mkdir(parents=True, exist_ok=True)
    CA_BUNDLE.write_text(Path(certifi.where()).read_text() + "\n" + "\n".join(pem_texts))
    return CA_BUNDLE


def connect_gigachat():
    if not KEY_FILE.exists():
        raise SystemExit(f"Нет файла {KEY_FILE.name} с ключом авторизации GigaChat API.")
    return GigaChat(
        credentials=KEY_FILE.read_text().strip(),
        scope="GIGACHAT_API_PERS",                  # доступ для физических лиц
        model=MODEL,
        ca_bundle_file=str(make_ca_bundle()),
        timeout=120,
    )


# =====================================================================
# Проверка чисел в тезисах
# =====================================================================
DATES = re.compile(r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b|\b\d{1,2}\.\d{4}\b")     # 30.09.2026, 08.2026
TERM_CODES = re.compile(r"\b\d+[WMY]\b")                                  # сроки ROISfix: 1W, 3M, 2Y
THOUSANDS = re.compile(r"(?<=\d)[   ](?=\d{3}\b)")                # 3 732,7 -> 3732,7
YEARS = re.compile(r"\b(?:19|20)\d{2}\b")                                   # 2026
NUMBER = re.compile(r"\d+(?:[.,]\d+)?")


def numbers_in_text(text):
    """Числа из текста (без дат, годов и кодов сроков): список (как написано, значение, знаков после запятой)"""
    cleaned = DATES.sub(" ", text)
    cleaned = TERM_CODES.sub(" ", cleaned)
    cleaned = THOUSANDS.sub("", cleaned)
    cleaned = YEARS.sub(" ", cleaned)
    found = []
    for raw in NUMBER.findall(cleaned):
        normalized = raw.replace(",", ".")
        decimals = len(normalized.split(".")[1]) if "." in normalized else 0
        found.append((raw, float(normalized), decimals))
    return found


def numbers_in_data(data):
    """Все числа из ответа функции (по модулю: «снизилась на 100 б.п.» = −100 в данных)"""
    if isinstance(data, bool):
        return set()
    if isinstance(data, (int, float)):
        return {abs(float(data))}
    if isinstance(data, dict):
        return set().union(*[numbers_in_data(value) for value in data.values()])
    if isinstance(data, list):
        return set().union(*[numbers_in_data(value) for value in data])
    return set()


def unconfirmed_numbers(text, data_numbers):
    """Числа из текста, которых нет в данных (с учётом округления, с которым они написаны)"""
    return [
        raw for raw, value, decimals in numbers_in_text(text)
        if not any(abs(value - known) <= 0.5 * 10 ** (-decimals) + 1e-9 for known in data_numbers)
    ]


def values_by_key(data, key):
    """Все значения поля key в ответе функции (на любой глубине)"""
    if isinstance(data, dict):
        found = {data[key]} if key in data else set()
        return found.union(*[values_by_key(value, key) for value in data.values()])
    if isinstance(data, list):
        return set().union(*[values_by_key(value, key) for value in data])
    return set()


def check_theses(text, tool_results):
    """Замечания к тезисам: числа, которых нет в данных, и противоречия состоянию ликвидности"""
    data_numbers = numbers_in_data(tool_results)
    problems = [f"числа {number} нет в данных" for number in unconfirmed_numbers(text, data_numbers)]

    liquidity_states = values_by_key(tool_results, "Состояние на отчётную дату")
    for state, opposite in [("дефицит", "профицит"), ("профицит", "дефицит")]:
        if liquidity_states == {state} and opposite in text.lower():
            problems.append(f"в данных {state} ликвидности, а в тексте — {opposite}")
    return problems


# =====================================================================
# Инструменты MCP -> функции GigaChat
# =====================================================================
def to_gigachat_function(tool):
    """Описание инструмента MCP в формате функции GigaChat"""
    schema = tool.input_schema or {}
    properties = {
        name: {key: value for key, value in spec.items() if key in {"type", "description", "enum"}}
        for name, spec in schema.get("properties", {}).items()
    }
    return Function(
        name=tool.name,
        description=tool.description or tool.name,
        parameters={"type": "object", "properties": properties, "required": schema.get("required", [])},
    )


# =====================================================================
# Диалог с моделью по одному разделу
# =====================================================================
async def ask_model(giga, session, functions, messages, tool_results, tool_calls, section):
    """Диалог до готового текста: модель вызывает функции, мы передаём ей результаты.
    Списки messages, tool_results и tool_calls пополняются по ходу диалога."""
    tokens = 0
    for _ in range(MAX_TOOL_CALLS + 1):
        response = await giga.achat(Chat(
            model=MODEL, messages=messages, functions=functions,
            function_call="auto", temperature=TEMPERATURE))
        tokens += response.usage.total_tokens
        choice = response.choices[0]
        messages.append(choice.message)

        if choice.finish_reason != "function_call" or choice.message.function_call is None:
            return choice.message.content, tokens            # модель написала текст

        name = choice.message.function_call.name
        arguments = choice.message.function_call.arguments or {}
        print(f"    вызов: {name}({json.dumps(arguments, ensure_ascii=False)})")

        result = await session.call_tool(name, arguments)
        result_text = result.content[0].text if result.content else "{}"
        try:
            tool_results.append(json.loads(result_text))
        except json.JSONDecodeError:
            result_text = json.dumps({"error": result_text}, ensure_ascii=False)

        tool_calls.append({"Раздел": section, "Функция": name,
                           "Аргументы": json.dumps(arguments, ensure_ascii=False)})
        messages.append(Messages(role=MessagesRole.FUNCTION, name=name, content=result_text))

    return "Модель превысила лимит обращений к данным и не написала текст.", tokens


async def write_section(giga, session, functions, section, topics, tool_calls):
    """Тезисы по разделу: черновик, проверка чисел, при необходимости — исправление"""
    print(f"\n{section}")
    messages = [
        Messages(role=MessagesRole.SYSTEM, content=SYSTEM_PROMPT),
        Messages(role=MessagesRole.USER, content=(
            f"Отчётный месяц: {REPORT_MONTH}. Раздел: «{section}». Темы раздела: {topics}. "
            "Запроси нужные данные через функции и напиши тезисы.")),
    ]
    tool_results = []
    total_tokens = 0
    fixes = 0

    text, tokens = await ask_model(giga, session, functions, messages, tool_results, tool_calls, section)
    total_tokens += tokens
    problems = check_theses(text, tool_results)

    while problems and fixes < MAX_FIX_ATTEMPTS:
        fixes += 1
        print(f"    замечания: {'; '.join(problems)} — прошу исправить (попытка {fixes})")
        messages.append(Messages(role=MessagesRole.USER, content=(
            f"Проверка нашла ошибки в тезисах: {'; '.join(problems)}. "
            "Перепиши тезисы целиком: используй только числа из ответов функций, ничего не вычисляй сам, "
            "состояние ликвидности бери из поля «Состояние на отчётную дату». "
            "При необходимости запроси данные ещё раз.")))
        text, tokens = await ask_model(giga, session, functions, messages, tool_results, tool_calls, section)
        total_tokens += tokens
        problems = check_theses(text, tool_results)

    print("    проверка: " + ("OK" if not problems else "; ".join(problems)))
    return {
        "Раздел": section,
        "Тезисы": text.strip(),
        "Проверка": "OK" if not problems else "ЕСТЬ ЗАМЕЧАНИЯ",
        "Замечания": "; ".join(problems),
        "Исправлений": fixes,
        "Токенов": total_tokens,
    }


# =====================================================================
# Запись результата
# =====================================================================
def save_to_excel(results, tool_calls):
    about = pd.DataFrame({
        "Параметр": ["Отчётный месяц", "Модель", "Создано", "Всего токенов"],
        "Значение": [REPORT_MONTH, MODEL, datetime.now().strftime("%d.%m.%Y %H:%M"),
                     sum(result["Токенов"] for result in results)],
    })
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        pd.DataFrame(results).to_excel(writer, sheet_name="Тезисы", index=False)
        pd.DataFrame(tool_calls).to_excel(writer, sheet_name="Вызовы функций", index=False)
        about.to_excel(writer, sheet_name="О запуске", index=False)

        sheet = writer.sheets["Тезисы"]
        for column, width in zip("ABCDEF", [24, 100, 30, 25, 13, 10]):
            sheet.column_dimensions[column].width = width
        for row in sheet.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
        writer.sheets["Вызовы функций"].column_dimensions["B"].width = 24
        writer.sheets["Вызовы функций"].column_dimensions["C"].width = 70
        writer.sheets["О запуске"].column_dimensions["A"].width = 18
        writer.sheets["О запуске"].column_dimensions["B"].width = 20


# =====================================================================
# Основной сценарий
# =====================================================================
def check_connection():
    giga = connect_gigachat()
    response = giga.chat("Ответь одной фразой: подключение работает?")
    print("Ответ GigaChat:", response.choices[0].message.content)
    print("Модель:", response.model, "| токенов:", response.usage.total_tokens)


async def main():
    giga = connect_gigachat()
    server = StdioServerParameters(command=sys.executable, args=[str(PROJECT_DIR / "mcp_server.py")],
                                   cwd=str(PROJECT_DIR))
    tool_calls = []

    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = (await session.list_tools()).tools
            functions = [to_gigachat_function(tool) for tool in tools]
            print(f"MCP-сервер запущен, инструментов: {len(tools)}: {', '.join(tool.name for tool in tools)}")

            results = [
                await write_section(giga, session, functions, section, topics, tool_calls)
                for section, topics in SECTIONS.items()
            ]

    save_to_excel(results, tool_calls)
    print(f"\nТезисы записаны: {OUTPUT_FILE}")


if __name__ == "__main__":
    if CHECK_ONLY:
        check_connection()
    elif REPORT_MONTH is None:
        raise SystemExit("Укажи отчётный месяц, например: python ai_commentary.py 2026-09")
    else:
        asyncio.run(main())
