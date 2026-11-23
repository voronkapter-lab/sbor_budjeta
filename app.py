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
    if not isinstance(value, str):
        return False
    text = re.sub(r"\s+", " ", value.strip())
    if not text or "вакансия" in text.casefold():
        return False
    words = re.findall(r"[А-Яа-яЁёA-Za-z-]+", text)
    return len(words) >= 3


def has_user_data(values):
    for value in values:
        if isinstance(value, str) and value.startswith("="):
            continue
