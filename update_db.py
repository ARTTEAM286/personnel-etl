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

file_path = r"C:\Users\Admin\Desktop\ДОВІРЕНА\ШПС_253ОШП_з_формулами.xlsm"

# ============================================
# READ EXCEL
# ============================================

df = pd.read_excel(
    file_path,
    sheet_name="ПЕРЕМІЩЕННЯ",
    header=1,
    usecols="B:M",
    engine="openpyxl"
)

# ============================================
# ORIGINAL ROW COUNT
# ============================================

original_rows = len(df)

# ============================================
# RENAME COLUMNS
# ============================================

df = df.rename(columns={

    'ПІБ': 'full_name',

    'Звання': 'military_rank',

    'ІПН': 'tax_number',

    'Підстава': 'reason',

    'Наказ (Дата)': 'order_date',

    'Наказ (Дата, Чий і Номер)': 'order_info',

    'Займана посада': 'current_position',

    'Нова посада': 'new_position',

    'Добовий (Дата)': 'daily_order_date',

    'Добовий (Номер)': 'daily_order_number',

    'Посаду прийняв': 'accepted_date',

    'Коментар': 'comment'
})

# ============================================
# REMOVE EMPTY ROWS
# ============================================

df = df.dropna(
    subset=["full_name"]
)

# ============================================
# REMOVE SERVICE ROW
# ============================================

df = df[
    df["full_name"] != "9"
]

# ============================================
# STRIP SPACES
# ============================================

text_cols = df.select_dtypes(
    include="object"
).columns

for col in text_cols:

    df[col] = (
        df[col]
        .astype(str)
        .str.strip()
        .replace("nan", pd.NA)
    )

# ============================================
# CONVERT DATES
# ============================================

date_columns = [

    "order_date",

    "daily_order_date",

    "accepted_date"
]

for col in date_columns:

    df[col] = pd.to_datetime(
        df[col],
        format="%d.%m.%Y",
        errors="coerce"
    )

# ============================================
# CLEAN TAX NUMBER
# ============================================

df["tax_number"] = (

    df["tax_number"]

    .astype(str)

    .str.replace(".0", "", regex=False)

    .replace("nan", pd.NA)
)

# ============================================
# REPLACE FAKE ZEROS
# ============================================

text_columns = [

    "current_position",

    "new_position",

    "daily_order_number"
]

for col in text_columns:

    df[col] = df[col].replace(
        0,
        pd.NA
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
# DEBUG
# ============================================

print("\n========== DATA PREVIEW ==========\n")

print(df.head())

print("\n========== DATA INFO ==========\n")

print(df.info())

# ============================================
# POSTGRESQL CONNECTION
# ============================================

engine = create_engine(
    "postgresql+psycopg2://postgres:3840@127.0.0.1:5432/arey_db"
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
                "SELECT COUNT(*) FROM personnel_movements"
            )
        ).scalar()

        # ====================================
        # CLEAR TABLE
        # ====================================

        conn.execute(
            text(
                "TRUNCATE TABLE personnel_movements"
            )
        )

        print("\nТаблиця очищена")

        # ====================================
        # UPLOAD NEW DATA
        # ====================================

        df.to_sql(
            "personnel_movements",
            conn,
            if_exists="append",
            index=False
        )

    # ========================================
    # SUCCESS SUMMARY
    # ========================================

    print("\n===================================")

    print("ОНОВЛЕННЯ БАЗИ ЗАВЕРШЕНО УСПІШНО")

    print("-----------------------------------")

    print(f"Було рядків у БД: {old_rows}")

    print(f"Очищено рядків у БД: {old_rows}")

    print(f"Було рядків в Excel: {original_rows}")

    print(f"Видалено сміттєвих рядків: {cleaned_rows}")

    print(f"Завантажено нових рядків: {final_rows}")

    print(f"Таблиця: personnel_movements")

    print("===================================\n")

except Exception as e:

    # ========================================
    # ERROR
    # ========================================

    print("\n===================================")

    print(f"ПОМИЛКА: {e}")

    print("===================================\n")