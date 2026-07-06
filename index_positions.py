# ============================================
# IMPORTS
# ============================================

import pandas as pd

from sqlalchemy import (
    create_engine,
    text
)

# ============================================
# FILE PATH
# ============================================

file_path = r"E:\ДОВІРЕНА\ШПС_253ОШП_з_формулами.xlsm"

# ============================================
# EXCEL SHEET
# ============================================

sheet_name = "ПІП"

# ============================================
# READ EXCEL
# ============================================

print("\n===================================")
print("ЧИТАННЯ EXCEL ФАЙЛУ")
print("===================================\n")

df = pd.read_excel(
    file_path,
    sheet_name=sheet_name,
    header=2,
    engine="openpyxl"
)

print(f"Зчитано рядків з Excel: {len(df)}")
print(f"Зчитано колонок: {len(df.columns)}")

# ============================================
# CONVERT COLUMN NAMES TO STRING
# ============================================

df.columns = df.columns.map(str)

# ============================================
# DEBUG ORIGINAL COLUMNS
# ============================================

print("\n===================================")
print("ОРИГІНАЛЬНІ НАЗВИ КОЛОНОК")
print("===================================\n")

for col in df.columns:
    print(col)

# ============================================
# COLUMN MAPPING
# ============================================

column_mapping = {

    "900": "type_key",
    "901": "unit_key",
    "902": "sequence_number",
    "109": "position_index",
    "903": "position_type",
    "904": "unit_level_1",
    "905": "unit_level_2",
    "906": "unit_level_3",
    "919": "unit_level_4",
    "110": "position_name",
    "111": "shpk",
    "112": "vos",
    "113": "tariff",
    "114": "shtat_number",
    "915": "position_enter_date",
    "916": "position_remove_date",
    "907_1": "code_main",
    "907": "code_duplication",
    "908": "code_1",
    "909": "code_2",
    "910": "code_3",
    "918": "code_4",
    "907_2": "subordinate_code",
    "918_1": "structure_position_code",
    "918_2": "main_position_code",
    "918_3": "combined_position_code",
    "918_4": "position_order_number",
    "911": "in_shtat",
    "912": "position_name_genitive",
    "904_1": "unit_level_4_genitive",
    "905_1": "unit_level_3_genitive",
    "906_1": "unit_level_2_genitive",
    "919_1": "unit_level_1_genitive",
    "912_1": "full_position_name_genitive",
    "917": "full_unit_name",
    "917_1": "unit_short_name"
}

# ============================================
# CHECK REQUIRED COLUMNS
# ============================================

print("\n===================================")
print("ПЕРЕВІРКА НАЯВНОСТІ КОЛОНОК")
print("===================================\n")

missing_columns = [

    col for col in column_mapping.keys()

    if col not in df.columns
]

if missing_columns:

    print("ВІДСУТНІ КОЛОНКИ:\n")

    for col in missing_columns:
        print(col)

    raise Exception(
        "Не всі колонки знайдені в Excel"
    )

print("Усі необхідні колонки знайдені")

# ============================================
# KEEP ONLY NEEDED COLUMNS
# ============================================

print("\n===================================")
print("ФІЛЬТРАЦІЯ КОЛОНОК")
print("===================================\n")

df = df[
    list(column_mapping.keys())
]

print(f"Залишено колонок: {len(df.columns)}")

# ============================================
# RENAME COLUMNS
# ============================================

df = df.rename(
    columns=column_mapping
)

print("Колонки успішно перейменовані")

# ============================================
# ORIGINAL ROW COUNT
# ============================================

original_rows = len(df)

# ============================================
# REMOVE EMPTY ROWS
# ============================================

print("\n===================================")
print("ОЧИЩЕННЯ ПОРОЖНІХ РЯДКІВ")
print("===================================\n")

df = df.dropna(
    subset=["position_index"]
)

df = df[
    df["position_index"]
    .astype(str)
    .str.strip() != ""
]

# ============================================
# TEXT COLUMNS
# ============================================

text_columns = [

    col for col in df.columns

    if col not in [

        "position_enter_date",
        "position_remove_date",
        "in_shtat"
    ]
]

# ============================================
# CLEAN TEXT COLUMNS
# ============================================

print("\n===================================")
print("ОЧИЩЕННЯ ТЕКСТОВИХ ПОЛІВ")
print("===================================\n")

for col in text_columns:

    df[col] = (

        df[col]

        .astype(str)

        .str.strip()

        .replace("nan", pd.NA)
    )

print(f"Очищено текстових колонок: {len(text_columns)}")

# ============================================
# DATE COLUMNS
# ============================================

date_columns = [

    "position_enter_date",
    "position_remove_date"
]

# ============================================
# CONVERT DATES
# ============================================

print("\n===================================")
print("КОНВЕРТАЦІЯ ДАТ")
print("===================================\n")

for col in date_columns:

    df[col] = pd.to_datetime(
        df[col],
        errors="coerce"
    )

print(f"Конвертовано колонок дат: {len(date_columns)}")

# ============================================
# BOOLEAN
# ============================================

print("\n===================================")
print("ОБРОБКА BOOLEAN")
print("===================================\n")

df["in_shtat"] = (

    df["in_shtat"]

    .astype(str)

    .str.upper()

    .map({

        "TRUE": True,
        "FALSE": False
    })
)

# ============================================
# FINAL ROW COUNT
# ============================================

final_rows = len(df)

# ============================================
# CLEANED ROW COUNT
# ============================================

cleaned_rows = original_rows - final_rows

# ============================================
# DEBUG DATAFRAME
# ============================================

print("\n===================================")
print("ПЕРЕВІРКА DATAFRAME")
print("===================================\n")

print(df.head())

print("\n===================================")
print("ІНФОРМАЦІЯ ПРО DATAFRAME")
print("===================================\n")

print(df.info())

# ============================================
# POSTGRESQL CONNECTION
# ============================================

print("\n===================================")
print("ПІДКЛЮЧЕННЯ ДО POSTGRESQL")
print("===================================\n")

engine = create_engine(
    "postgresql+psycopg2://postgres:postgres@127.0.0.1:5432/arey_db"
)

# ============================================
# DATABASE UPDATE
# ============================================

try:

    with engine.begin() as conn:

        # ====================================
        # GET OLD ROW COUNT
        # ====================================

        old_rows = conn.execute(
            text(
                "SELECT COUNT(*) FROM unit_positions"
            )
        ).scalar()

        # ====================================
        # CLEAR TABLE
        # ====================================

        conn.execute(
            text(
                "TRUNCATE TABLE unit_positions"
            )
        )

        print("Таблиця unit_positions очищена")

        # ====================================
        # UPLOAD DATA
        # ====================================

        print("\nПочаток завантаження даних...\n")

        df.to_sql(
            "unit_positions",
            conn,
            if_exists="append",
            index=False
        )

        print("Дані успішно завантажені")

    # ========================================
    # SUCCESS SUMMARY
    # ========================================

    print("\n===================================")
    print("ІМПОРТ ЗАВЕРШЕНО УСПІШНО")
    print("===================================\n")

    print(f"Було рядків у БД: {old_rows}")

    print(f"Очищено рядків у БД: {old_rows}")

    print(f"Було рядків в Excel: {original_rows}")

    print(f"Видалено порожніх рядків: {cleaned_rows}")

    print(f"Завантажено нових рядків: {final_rows}")

    print(f"Таблиця PostgreSQL: unit_positions")

    print("\n===================================\n")

except Exception as e:

    # ========================================
    # ERROR
    # ========================================

    print("\n===================================")
    print("ПОМИЛКА ПІД ЧАС ІМПОРТУ")
    print("===================================\n")

    print(e)

    print("\n===================================\n")