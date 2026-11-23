import io
import math
import re
import traceback
from collections import defaultdict

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# -------------------- НАСТРОЙКИ --------------------
st.set_page_config(page_title="Сборка бюджета", page_icon="📊", layout="wide")

st.title("📊 Сборка бюджета")
st.caption("Загрузка исходников → проверка → сборка единого Excel.")
st.info(
    "Файл формируется даже при наличии ошибок. Проблемы и предупреждения "
    "попадут на отдельную вкладку «Журнал ошибок»."
)

# Загрузчик показывается сразу, до объявления обработчиков.
uploaded_files = st.file_uploader(
    "Загрузите Excel-файлы проектов",
    type=["xlsx"],
    accept_multiple_files=True,
    help="В каждом файле должны быть листы «Титульная», «Команда» и «Закупки»."
)

selected_count = len(uploaded_files) if uploaded_files else 0
if selected_count:
    st.success(f"Выбрано файлов: {selected_count}")
else:
    st.warning("Сначала выберите один или несколько файлов .xlsx.")

build_clicked = st.button(
    "▶ Собрать бюджет",
    type="primary",
    use_container_width=True,
    disabled=(selected_count == 0),
)

TITLE_SHEET = "Титульная"
TEAM_SHEET = "Команда"
PURCHASE_SHEET = "Закупки"

TITLE_LABELS = [
    "Идентификатор проекта",
    "Тип проекта",
    "Название проекта",
    "Руководитель",
    "Площадки",
    "Публикаций в 2027 году",
    "РИД в 2027 году",
]

TEAM_HEADERS = [
    "Сотрудник", "Роль", "Доля ставки", "Рабочих дней в месяце",
    "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12",
    "Трудозатраты, месяцы", "Доля рабочего времени на проекте",
    "Project_ID", "Filename", "Тип проекта", "Название проекта",
]

PURCHASE_HEADERS = [
    "ОЦДИ", "Категория закупок", "Исполнители внутренних услуг",
    "Что закупаем", "Цена", "Количество", "Сумма, руб.",
    "Дата поставки", "ФЛС", "Потрачено", "ZAKUP", "Комментарий",
    "Project_ID", "Filename", "Тип проекта", "Название проекта",
    "Project NameID",
]

TITLE_HEADERS = [
    "Идентификатор проекта", "Тип проекта", "Название проекта", "Руководитель",
    "Площадки", "Публикаций в 2027 году", "РИД в 2027 году",
    "Бюджет на закупки", "Потрачено", "ФОТ", "Внутренние услуги",
    "Магазинчик", "Оборудование", "Расходники", "Услуги",
    "Публикации", "Командировки", "Оборудование (ОЦДИ)", "Filename",
]

LOG_HEADERS = [
    "Уровень", "Filename", "Project_ID", "Лист", "Строка / ячейка",
    "Сотрудник или закупка", "Проверка", "Описание проблемы",
    "Исходное значение", "Что удалось рассчитать", "Рекомендация",
]

KNOWN_CATEGORIES = {
    "Расходники", "Оборудование", "Магазинчик", "Услуги",
    "Внутренние услуги", "Командировки", "Публикации",
}


# -------------------- ОБЩИЕ ФУНКЦИИ --------------------
def clean(value):
    if value is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value


def to_number(value):
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value) if math.isfinite(value) else None
    if isinstance(value, str):
        text = value.strip().replace(" ", "").replace(",", ".")
        try:
            result = float(text)
            return result if math.isfinite(result) else None
        except ValueError:
            return None
    return None


def add_issue(logs, level, filename, project_id, sheet, cell, subject,
              check, description, original=None, calculated=None,
              recommendation="Проверить исходный файл."):
    logs.append([
        level, filename, project_id, sheet, cell, subject, check,
        description, original, calculated, recommendation
    ])


def is_full_person_name(value):
    """True only for a plain three-part personal name, not roles or notes."""
    if not isinstance(value, str):
        return False
    text = re.sub(r"\\s+", " ", value.strip())
    if not text:
        return False

    # Notes, dates, brackets, vacancy labels and explanatory suffixes are not FIO.
    if re.search(r"[()\\[\\]{}0-9*/]", text):
        return False

    words = text.split(" ")
    if len(words) != 3:
        return False

    role_words = {
        "лаборант", "лаборанта", "лаборанты", "техник", "инженер",
        "ассистент", "специалист", "руководитель", "менеджер",
        "стажер", "стажёр", "вакансия", "вакантно", "должность",
        "уборщик", "оператор", "координатор", "администратор",
    }
    if any(word.casefold().strip(".,;:") in role_words for word in words):
        return False

    # Each component must be a name-like alphabetic token (hyphenated names allowed).
    return all(re.fullmatch(r"[А-Яа-яЁёA-Za-z]+(?:-[А-Яа-яЁёA-Za-z]+)*", word) for word in words)


def has_user_data(values):
    for value in values:
        if isinstance(value, str) and value.startswith("="):
            continue
        if clean(value) is not None:
            return True
    return False


# -------------------- ЧТЕНИЕ ИСХОДНОГО ФАЙЛА --------------------
def read_project(upload, logs):
    filename = upload.name
    file_bytes = upload.getvalue()

    try:
        # Формулы читаем как формулы: D/Q/R и G пересчитываем сами.
        workbook = load_workbook(io.BytesIO(file_bytes), data_only=False, read_only=False)
    except Exception as exc:
        add_issue(logs, "Ошибка", filename, None, "-", "-", "Открытие файла",
                  "Не удалось прочитать Excel-файл.", repr(exc), None,
                  "Пересохранить исходник в формате .xlsx и загрузить повторно.")
        return None

    required_sheets = [TITLE_SHEET, TEAM_SHEET, PURCHASE_SHEET]
    missing = [name for name in required_sheets if name not in workbook.sheetnames]
    for name in missing:
        add_issue(logs, "Ошибка", filename, None, name, "-", "Обязательный лист",
                  "В книге отсутствует обязательный лист.", name, None)
    if missing:
        return None

    # Титульная: используем только значения B1:B7. Названия строк A1:A7
    # проверяем, но не используем как источник значений.
    title_ws = workbook[TITLE_SHEET]
    title_values = [clean(title_ws.cell(row, 2).value) for row in range(1, 8)]
    project_id = title_values[0]
    for row_num, value in enumerate(title_values, start=1):
        expected = TITLE_LABELS[row_num - 1]
        actual = clean(title_ws.cell(row_num, 1).value)
        if actual != expected:
            actual_text = "пусто" if actual is None else str(actual)
            description = (
                f"В A{row_num} ожидалось «{expected}», найдено «{actual_text}». "
                f"Значение перенесено по позиции из B{row_num}."
            )
            add_issue(logs, "Предупреждение", filename, project_id, TITLE_SHEET,
                      f"A{row_num}", "Название показателя", description,
                      actual, value, "Проверить, не изменился ли смысл строки.")
        if value is None:
            level = "Ошибка" if row_num in (1, 3) else "Предупреждение"
            add_issue(logs, level, filename, project_id, TITLE_SHEET,
                      f"B{row_num}", "Пустое значение шапки",
                      f"Не заполнено значение «{expected}».", None, None,
                      "Проверить и при необходимости заполнить исходную шапку.")

    project = {
        "filename": filename,
        "project_id": project_id,
        "project_type": title_values[1],
        "project_name": title_values[2],
        "title": title_values,
        "team": [],
        "purchases": [],
    }

    # Команда: A:R. Строки с пользовательскими данными сохраняем,
    # включая вакансии и текстовые пояснения.
    team_ws = workbook[TEAM_SHEET]
    expected_team = [
        "Сотрудник", "Роль", "Доля ставки", "Рабочих дней в месяце",
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12,
        "Трудозатраты, месяцы", "Доля рабочего времени на проекте"
    ]
    actual_team = [team_ws.cell(1, col).value for col in range(1, 19)]
    if actual_team != expected_team:
        add_issue(logs, "Предупреждение", filename, project_id, TEAM_SHEET,
                  "A1:R1", "Заголовки команды",
                  "Заголовки отличаются от эталона. Данные читаются по позициям A:R.",
                  str(actual_team), None, "Проверить изменения структуры листа.")

    for row_num in range(2, team_ws.max_row + 1):
        values = [team_ws.cell(row_num, col).value for col in range(1, 19)]
        if not has_user_data(values):
            continue

        employee = clean(values[0])
        rate = to_number(values[2])
        monthly = values[4:16]
        working_days = math.floor(21 * rate + 0.5) if rate is not None and rate >= 0 else None

        numeric_days = []
        unusual_cells = []
        for month_index, value in enumerate(monthly, start=5):
            if clean(value) is None or (isinstance(value, str) and value.strip() == "-"):
                continue
            number = to_number(value)
            if number is None:
                unusual_cells.append((month_index, value))
            else:
                numeric_days.append(number)

        # Нечисловые значения (например, 4*) не трактуем как ноль:
        # Q/R не рассчитываем, чтобы не выдавать неполный результат за точный.
        if unusual_cells:
            effort_months = None
            share = None
        else:
            total_days = sum(numeric_days)
            effort_months = total_days / 21
            dash_count = sum(
                1 for value in monthly
                if isinstance(value, str) and value.strip() == "-"
            )
            denominator = (12 - dash_count) * working_days if working_days else None
            share = total_days / denominator if denominator else None

        output_row = (
            values[:3] + [working_days] + monthly +
            [effort_months, share, project_id, filename,
             project["project_type"], project["project_name"]]
        )
        project["team"].append(output_row)

        if rate is None and clean(values[2]) is not None:
            add_issue(logs, "Ошибка", filename, project_id, TEAM_SHEET,
                      f"C{row_num}", employee or "Строка команды",
                      "Доля ставки не является числом.", values[2], None)

        if clean(values[0]) is None and any(clean(v) is not None for v in values[1:]):
            add_issue(logs, "Предупреждение", filename, project_id, TEAM_SHEET,
                      f"A{row_num}", "Строка без ФИО",
                      "Строка сохранена в команде, но не содержит ФИО.",
                      None, None, "Проверить, является ли строка пояснением или неполной записью.")

        for col_num, original in unusual_cells:
            add_issue(logs, "Предупреждение", filename, project_id, TEAM_SHEET,
                      f"{get_column_letter(col_num)}{row_num}", employee or "Строка команды",
                      "Нечисловая помесячная загрузка",
                      "Значение сохранено как есть; трудозатраты и доля по этой строке не рассчитаны.",
                      original, None, "Уточнить значение. После уточнения пересчитать Q и R.")

    # Закупки: строго сохраняем согласованную структуру A:L.
    # Дополнительные столбцы источника не подмешиваем в выходные 17 колонок.
    purchase_ws = workbook[PURCHASE_SHEET]
    expected_purchase = [
        "ОЦДИ", "Категория закупок", "Исполнители внутренних услуг",
        "Что закупаем", "Цена", "Количество", "Сумма, руб.",
        "Дата поставки", "ФЛС", "Потрачено", "ZAKUP", "Комментарий"
    ]
    actual_purchase = [purchase_ws.cell(1, col).value for col in range(1, 13)]
    if actual_purchase != expected_purchase:
        add_issue(logs, "Предупреждение", filename, project_id, PURCHASE_SHEET,
                  "A1:L1", "Заголовки закупок",
                  "Заголовки отличаются от эталона. Данные читаются по позициям A:L.",
                  str(actual_purchase), None, "Проверить изменения структуры листа.")

    for row_num in range(2, purchase_ws.max_row + 1):
        values = [purchase_ws.cell(row_num, col).value for col in range(1, 13)]
        if not has_user_data(values):
            continue

        category = clean(values[1])
        executor = clean(values[2])
        item = clean(values[3])
        price = to_number(values[4])
        quantity = to_number(values[5])
        source_sum_raw = values[6]
        source_sum = to_number(source_sum_raw)
        calculated_sum = price * quantity if price is not None and quantity is not None else None
        ocdi = "да" if category == "Оборудование" and price is not None and price > 500000 else "нет"
        subject = item or category or f"Строка {row_num}"

        if category is None:
            add_issue(logs, "Ошибка", filename, project_id, PURCHASE_SHEET,
                      f"B{row_num}", subject, "Не заполнена обязательная категория закупок.",
                      None, None)
        elif category not in KNOWN_CATEGORIES:
            add_issue(logs, "Предупреждение", filename, project_id, PURCHASE_SHEET,
                      f"B{row_num}", subject, "Неизвестная категория закупок.",
                      category, category, "Категория сохранена; проверить классификацию.")

        if category == "Внутренние услуги" and executor is None:
            add_issue(logs, "Ошибка", filename, project_id, PURCHASE_SHEET,
                      f"C{row_num}", subject,
                      "Для категории «Внутренние услуги» не указан исполнитель.", None, None)

        if item is None:
            add_issue(logs, "Ошибка", filename, project_id, PURCHASE_SHEET,
                      f"D{row_num}", subject, "Не заполнено поле «Что закупаем».",
                      None, calculated_sum)
        if price is None:
            add_issue(logs, "Ошибка", filename, project_id, PURCHASE_SHEET,
                      f"E{row_num}", subject, "Цена отсутствует или не является числом.",
                      values[4], calculated_sum)
        if quantity is None:
            add_issue(logs, "Ошибка", filename, project_id, PURCHASE_SHEET,
                      f"F{row_num}", subject, "Количество отсутствует или не является числом.",
                      values[5], calculated_sum)

        # В исходниках G обычно содержит формулу =Eстрока*Fстрока.
        # Это ожидаемая формула, а не отсутствие суммы. Если там число — сверяем.
        if isinstance(source_sum_raw, str) and source_sum_raw.startswith("="):
            normalized_formula = source_sum_raw.replace("$", "").replace(" ", "").upper()
            expected_formula = f"=E{row_num}*F{row_num}".upper()
            if normalized_formula != expected_formula:
                add_issue(logs, "Предупреждение", filename, project_id, PURCHASE_SHEET,
                          f"G{row_num}", subject,
                          "Формула исходной суммы отличается от ожидаемой E×F.",
                          source_sum_raw, calculated_sum,
                          "Проверить формулу исходника; в итог записан собственный расчёт.")
        elif source_sum is None:
            add_issue(logs, "Ошибка", filename, project_id, PURCHASE_SHEET,
                      f"G{row_num}", subject, "Исходная сумма отсутствует или не является числом.",
                      source_sum_raw, calculated_sum,
                      "В итог записана сумма E×F, если оба значения доступны.")
        elif calculated_sum is not None and not math.isclose(
            source_sum, calculated_sum, rel_tol=1e-8, abs_tol=0.01
        ):
            add_issue(logs, "Ошибка", filename, project_id, PURCHASE_SHEET,
                      f"G{row_num}", subject, "Исходная сумма не совпадает с ценой × количеством.",
                      source_sum, calculated_sum,
                      "В итог записана пересчитанная сумма.")

        project_name_id = (
            f"{project_id} {project['project_name'] or ''}".strip()
            if project_id else project["project_name"]
        )
        project["purchases"].append([
            ocdi, category, executor, item, values[4], values[5], calculated_sum,
            values[7], values[8], values[9], values[10], values[11],
            project_id, filename, project["project_type"], project["project_name"],
            project_name_id
        ])

    return project


# -------------------- ФОРМИРОВАНИЕ ИТОГОВОГО EXCEL --------------------
def build_output(projects, logs):
    title_rows, team_rows, purchase_rows = [], [], []
    seen_ids = set()

    for project in projects:
        project_id = project["project_id"]
        if project_id and project_id in seen_ids:
            add_issue(logs, "Ошибка", project["filename"], project_id, TITLE_SHEET,
                      "B1", "Повтор ID проекта",
                      "Такой Project_ID уже встречался в другой загруженной книге.",
                      project_id, None, "Проверить, не загружен ли дубликат проекта.")
        if project_id:
            seen_ids.add(project_id)

        totals = defaultdict(float)
        for row in project["purchases"]:
            category, amount, ocdi = row[1], row[6], row[0]
            if amount is None:
                continue
            if category == "Оборудование":
                target = "Оборудование (ОЦДИ)" if ocdi == "да" else "Оборудование"
                totals[target] += amount
            elif category:
                totals[category] += amount

        internal = totals["Внутренние услуги"]
        shop = totals["Магазинчик"]
        equipment = totals["Оборудование"]
        consumables = totals["Расходники"]
        services = totals["Услуги"]
        publications = totals["Публикации"]
        travel = totals["Командировки"]
        equipment_ocdi = totals["Оборудование (ОЦДИ)"]

        # В бюджет входят все рассчитанные суммы, кроме внутренних услуг,
        # магазинчика и оборудования с ОЦДИ.
        budget = sum(
            amount for category, amount in totals.items()
            if category not in {"Внутренние услуги", "Магазинчик", "Оборудование (ОЦДИ)"}
        )

        title_rows.append(
            project["title"] + [
                budget, None, None, internal, shop, equipment,
                consumables, services, publications, travel,
                equipment_ocdi, project["filename"]
            ]
        )
        team_rows.extend(project["team"])
        purchase_rows.extend(project["purchases"])

    # Суммируем R по одинаковым полноценным ФИО между всеми загруженными проектами.
    employee_totals = defaultdict(float)
    employee_files = defaultdict(list)
    employee_ids = defaultdict(list)
    for row in team_rows:
        name, share = row[0], row[17]
        if is_full_person_name(name) and isinstance(share, (int, float)):
            key = re.sub(r"\s+", " ", name.strip()).casefold()
            employee_totals[key] += share
            employee_files[key].append(row[19])
            employee_ids[key].append(row[18])

    for key, total in employee_totals.items():
        if total > 1 + 1e-9:
            name = next(row[0] for row in team_rows if is_full_person_name(row[0]) and
                        re.sub(r"\s+", " ", row[0].strip()).casefold() == key)
            files = ", ".join(dict.fromkeys(employee_files[key]))
            ids = ", ".join(dict.fromkeys(str(x) for x in employee_ids[key] if x))
            add_issue(logs, "Предупреждение", files, ids, TEAM_SHEET, "R",
                      name, "Суммарная доля рабочего времени по проектам превышает 1.",
                      total, total, "Уточнить распределение занятости сотрудника.")

    # Корректировка титула: полученные внутренние заказы уменьшают
    # бюджет закупок и расходники проекта-исполнителя.
    # Учитываем только категорию «Внутренние услуги» и заполненный столбец C.
    project_by_id = {
        str(project["project_id"]).strip().upper(): project
        for project in projects
        if project.get("project_id") is not None and str(project["project_id"]).strip()
    }
    received_internal_by_id = defaultdict(float)

    for project in projects:
        for purchase in project["purchases"]:
            category = clean(purchase[1])
            executor = clean(purchase[2])
            amount = to_number(purchase[6])

            if category != "Внутренние услуги" or executor is None or amount is None:
                continue

            executor_text = str(executor).strip()
            executor_id = executor_text[:6].upper()

            if not re.fullmatch(r"[A-ZА-ЯЁ0-9]{6}", executor_id):
                add_issue(
                    logs, "Предупреждение", project["filename"], project["project_id"],
                    PURCHASE_SHEET, "C", purchase[3] or "Внутренние услуги",
                    "Не удалось определить ID проекта-исполнителя по первым 6 символам.",
                    executor, amount,
                    "Проверить значение исполнителя в столбце C. Строка не включена в корректировку."
                )
                continue

            if executor_id not in project_by_id:
                add_issue(
                    logs, "Предупреждение", project["filename"], project["project_id"],
                    PURCHASE_SHEET, "C", purchase[3] or "Внутренние услуги",
                    "Проект-исполнитель не найден среди загруженных проектов.",
                    executor, amount,
                    "Загрузить книгу проекта-исполнителя или проверить ID. Сумма не включена в корректировку."
                )
                continue

            received_internal_by_id[executor_id] += amount

    corrected_title_rows = []
    for project, ordinary_row in zip(projects, title_rows):
        project_id = project.get("project_id")
        project_key = str(project_id).strip().upper() if project_id is not None else ""
        received = received_internal_by_id.get(project_key, 0.0)

        corrected_row = list(ordinary_row)
        # H — бюджет закупок; N — расходники.
        corrected_row[7] = (to_number(corrected_row[7]) or 0.0) - received
        corrected_row[13] = (to_number(corrected_row[13]) or 0.0) - received
        corrected_title_rows.append(corrected_row)

        if received:
            if corrected_row[13] < 0:
                add_issue(
                    logs, "Предупреждение", project["filename"], project_id,
                    "Титул с корректировкой", "N",
                    project.get("project_name") or project_id,
                    "Полученные внутренние заказы превышают бюджет расходников.",
                    received, corrected_row[13],
                    "Отрицательное значение сохранено по согласованному правилу; проверить план расходников."
                )
            if corrected_row[7] < 0:
                add_issue(
                    logs, "Предупреждение", project["filename"], project_id,
                    "Титул с корректировкой", "H",
                    project.get("project_name") or project_id,
                    "После вычитания внутренних заказов бюджет закупок стал отрицательным.",
                    received, corrected_row[7],
                    "Отрицательное значение сохранено по согласованному правилу; проверить бюджет проекта."
                )

    output = Workbook()
    output.remove(output.active)

    sheets = [
        ("Титульные", TITLE_HEADERS, title_rows),
        ("Титул с корректировкой", TITLE_HEADERS, corrected_title_rows),
        ("Команда", TEAM_HEADERS, team_rows),
        ("Закупки", PURCHASE_HEADERS, purchase_rows),
        ("Журнал ошибок", LOG_HEADERS, logs),
    ]

    for sheet_name, headers, rows in sheets:
        ws = output.create_sheet(sheet_name)
        ws.append(headers)
        for row in rows:
            ws.append(list(row))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 34
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="244062")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            sample = [
                len(str(ws.cell(r, col_idx).value or ""))
                for r in range(1, min(ws.max_row, 150) + 1)
            ]
            ws.column_dimensions[letter].width = min(42, max(12, max(sample, default=12) + 2))

    if not logs:
        ws = output["Журнал ошибок"]
        ws.append(["Нет ошибок", None, None, None, None, None,
                   "Проверка завершена", "Ошибок и предупреждений не обнаружено."])

    buffer = io.BytesIO()
    output.save(buffer)
    return buffer.getvalue()


# -------------------- ИНТЕРФЕЙС --------------------

if build_clicked and uploaded_files:
        issues = []
        projects = []
        progress = st.progress(0)
        status = st.empty()

        for index, uploaded in enumerate(uploaded_files, start=1):
            status.write(f"Обрабатывается {uploaded.name} ({index}/{len(uploaded_files)})")
            try:
                project = read_project(uploaded, issues)
                if project is not None:
                    projects.append(project)
            except Exception as exc:
                add_issue(
                    issues, "Ошибка", uploaded.name, None, "Обработка файла", "-",
                    "Непредвиденная ошибка", repr(exc), traceback.format_exc(),
                    None, "Файл пропущен; остальные файлы продолжают обрабатываться."
                )
            progress.progress(index / len(uploaded_files))

        try:
            excel_bytes = build_output(projects, issues)
            error_count = sum(1 for item in issues if item[0] == "Ошибка")
            warning_count = sum(1 for item in issues if item[0] == "Предупреждение")

            st.success(f"Сборка завершена: обработано проектов {len(projects)} из {len(uploaded_files)}.")
            col1, col2, col3 = st.columns(3)
            col1.metric("Проектов", len(projects))
            col2.metric("Ошибок", error_count)
            col3.metric("Предупреждений", warning_count)

            st.download_button(
                "⬇ Скачать итоговый Excel",
                data=excel_bytes,
                file_name="Итоговая_сборка_бюджета.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            if issues:
                st.subheader("Журнал проверки")
                st.dataframe(
                    pd.DataFrame(issues, columns=LOG_HEADERS),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success("Ошибок и предупреждений не обнаружено.")
        except Exception as exc:
            st.error("Не удалось сформировать итоговый Excel.")
            st.code(traceback.format_exc())
