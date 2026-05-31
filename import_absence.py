import pandas as pd
import psycopg2
from io import StringIO
from datetime import datetime

# ============================================================
# НАЛАШТУВАННЯ
# ============================================================

EXCEL_FILE = r"E:\ДОВІРЕНА\ШПС_253ОШП_з_формулами.xlsm"
SHEET_NAME = "ТВ"

DB_CONFIG = {
    "host": "localhost",
    "port": "5432",
    "dbname": "arey_db",
    "user": "postgres",
    "password": "postgres"
}

# ============================================================
# ПОЧАТОК РОБОТИ
# ============================================================

start_time = datetime.now()

print("\n===================================")
print("ІМПОРТ ТИМЧАСОВО ВІДСУТНІХ")
print("===================================\n")

# ============================================================
# ЧИТАННЯ EXCEL
# ============================================================

print("Читання Excel...")

df = pd.read_excel(
    EXCEL_FILE,
    sheet_name=SHEET_NAME,
    header=2,          # третій рядок містить коди колонок
    dtype=str
)

print(f"Зчитано рядків: {len(df):,}")

# ============================================================
# ВІДПОВІДНІСТЬ КОДІВ EXCEL ДО КОЛОНОК БД
# ============================================================

column_mapping = {

    "500": "absence_order_number",
    "501": "departure_unit_number",
    "500.1": "departure_unit_order_number",
    "502": "tax_number",

    "65": "rank_name",
    "66": "full_name",
    "67": "position_index",

    "68": "departure_reason",
    "69": "departure_location",

    "70": "departure_date",
    "70.1": "deregistration_date",

    "71": "departure_order_date",
    "71.0": "departure_order_number",

    "72": "return_period_days",
    "73": "commander_confirmation_date",

    "74": "actual_arrival_date",
    "74.1": "registration_date",

    "75": "arrival_order_date",
    "75.0": "arrival_order_number",

    "76": "additional_info",

    "503": "absence_type",
    "504": "absence_short_name",

    "304": "status",
    "304_0": "concentration_area_status",

    "505": "arrived_flag",
    "506": "departure_document_registered",

    "507": "current_status",

    "508": "position_duplicate_count",

    "509": "departure_document_exists",
    "510": "open_report",
    "511": "checked_flag"
}

# ============================================================
# ПЕРЕВІРКА НАЯВНОСТІ КОЛОНОК
# ============================================================

print("\nПеревірка структури Excel...")

missing_columns = []

for col in column_mapping.keys():
    if col not in df.columns:
        missing_columns.append(col)

if missing_columns:

    print("\nВІДСУТНІ КОЛОНКИ:")

    for col in missing_columns:
        print(col)

    raise Exception("Не всі колонки знайдені в Excel")

print("Усі необхідні колонки знайдено.")

# ============================================================
# ЗАЛИШАЄМО ЛИШЕ ПОТРІБНІ КОЛОНКИ
# ============================================================

df = df[list(column_mapping.keys())]

# ============================================================
# ПЕРЕЙМЕНОВУЄМО КОЛОНКИ ПІД БД
# ============================================================

df.rename(
    columns=column_mapping,
    inplace=True
)

# ============================================================
# ОЧИЩЕННЯ ДАНИХ
# ============================================================

print("Очищення даних...")

for col in df.columns:

    df[col] = (
        df[col]
        .fillna("")
        .astype(str)
        .str.replace(r"[\r\n\t]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

print("Перевірка переносів рядків...")

bad_values = 0

for col in df.columns:

    mask = (
        df[col].astype(str).str.contains("\n", na=False)
        |
        df[col].astype(str).str.contains("\r", na=False)
    )

    bad_values += mask.sum()

print(f"Знайдено проблемних значень: {bad_values}")

# ============================================================
# ОБРОБКА ДАТ
# ============================================================

print("Обробка дат...")

date_columns = [
    "departure_date",
    "deregistration_date",
    "departure_order_date",
    "commander_confirmation_date",
    "actual_arrival_date",
    "registration_date",
    "arrival_order_date"
]

for col in date_columns:

    df[col] = pd.to_datetime(
    df[col],
    errors="coerce"
    ).dt.date

# ============================================================
# ВИДАЛЕННЯ ПОВНІСТЮ ПУСТИХ РЯДКІВ
# ============================================================

rows_before = len(df)

df = df[
    df["tax_number"].astype(str).str.strip() != ""
]

removed_rows = rows_before - len(df)

print(f"Видалено порожніх рядків: {removed_rows}")

# ============================================================
# ПІДКЛЮЧЕННЯ ДО POSTGRESQL
# ============================================================

print("\nПідключення до PostgreSQL...")

conn = psycopg2.connect(**DB_CONFIG)
cur = conn.cursor()

# ============================================================
# КІЛЬКІСТЬ РЯДКІВ ДО ОЧИЩЕННЯ
# ============================================================

cur.execute("""
SELECT COUNT(*)
FROM temporary_absence
""")

old_rows = cur.fetchone()[0]

print(f"Було записів у БД: {old_rows}")


# ============================================================
# ОЧИЩЕННЯ ТАБЛИЦІ
# ============================================================

print("Очищення temporary_absence...")

cur.execute("""
TRUNCATE TABLE temporary_absence RESTART IDENTITY;
""")

conn.commit()

# ============================================================
# ПІДГОТОВКА ДО COPY
# ============================================================

print("Підготовка даних до завантаження...")

buffer = StringIO()

df.to_csv(
    buffer,
    sep="\t",
    header=False,
    index=False,
    na_rep=""
)

buffer.seek(0)

# ============================================================
# ЗАВАНТАЖЕННЯ В БАЗУ
# ============================================================

print("Завантаження даних у PostgreSQL...")

cur.copy_from(
    buffer,
    "temporary_absence",
    sep="\t",
    null="",
    columns=df.columns
)

conn.commit()

# ============================================================
# КОНТРОЛЬНИЙ ПІДРАХУНОК
# ============================================================

cur.execute("""
SELECT COUNT(*)
FROM temporary_absence
""")

loaded_rows = cur.fetchone()[0]

cur.close()
conn.close()

# ============================================================
# ПІДСУМОК
# ============================================================

execution_time = datetime.now() - start_time

print("\n===================================")
print("ІМПОРТ ЗАВЕРШЕНО УСПІШНО")
print("===================================\n")

print(f"Було рядків у БД: {old_rows}")
print(f"Очищено рядків у БД: {old_rows}")

print(f"Було рядків в Excel: {rows_before}")
print(f"Видалено порожніх рядків: {removed_rows}")

print(f"Завантажено нових рядків: {loaded_rows}")

print(f"Таблиця PostgreSQL: temporary_absence")

print(f"Час виконання: {execution_time}")

print("\n===================================\n")