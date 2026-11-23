import io
import math
import re
from collections import defaultdict

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Сборка бюджета", page_icon="📊", layout="wide")
st.title("📊 Сборка бюджета")
st.caption("Приложение запущено. Подготовка интерфейса…")

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
EXCLUDED_FROM_PURCHASE_BUDGET = {
    "ФОТ", "Внутренние услуги", "Магазинчик"
}

def clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        return s if s else None
    return v

def is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)

def num(v):
    if is_number(v):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return None
    return None

def add_log(logs, level, filename, project_id, sheet, cell, subject, check, description, original=None, calculated=None, recommendation="Проверить исходный файл."):
    logs.append([
        level, filename, project_id, sheet, cell, subject, check,
        description, original, calculated, recommendation
    ])

def is_formula_only_row(values):
    meaningful = []
    for v in values:
        if isinstance(v, str) and v.startswith("="):
            continue
        if clean(v) is not None:
            meaningful.append(v)
    return not meaningful

def read_project(uploaded_file, logs):
    filename = uploaded_file.name
    raw = uploaded_file.getvalue()
    try:
        wb = load_workbook(io.BytesIO(raw), data_only=False)
    except Exception as e:
        add_log(logs, "Ошибка", filename, None, "-", "-", "Чтение файла",
                "Файл Excel не прочитан", str(e), recommendation="Пересохранить файл в формате .xlsx.")
        return None

    for required in (TITLE_SHEET, TEAM_SHEET, PURCHASE_SHEET):
        if required not in wb.sheetnames:
            add_log(logs, "Ошибка", filename, None, required, "-", "Наличие листа",
                    f"Не найден обязательный лист «{required}».")
    if not all(s in wb.sheetnames for s in (TITLE_SHEET, TEAM_SHEET, PURCHASE_SHEET)):
        return None

    tws = wb[TITLE_SHEET]
    title_values = [clean(tws.cell(row=i, column=2).value) for i in range(1, 8)]
    project_id = title_values[0]
    for i, value in enumerate(title_values, start=1):
        expected = TITLE_LABELS[i-1]
        actual = clean(tws.cell(row=i, column=1).value)
        if actual != expected:
            actual_label = actual if actual is not None else "пусто"
            description = (
