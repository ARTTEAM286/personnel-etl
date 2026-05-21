# %%
import pandas as pd
from sqlalchemy import create_engine

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

df = df.dropna(subset=["full_name"])

# ============================================
# REMOVE SERVICE ROW
# ============================================

df = df[df["full_name"] != "9"]

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
    pd.to_numeric(
        df["tax_number"],
        errors="coerce"
    )
    .astype("Int64")
    .astype(str)
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
    df[col] = df[col].replace(0, pd.NA)

# ============================================
# DEBUG
# ============================================

print(df.head())

print(df.info())

# ============================================
# POSTGRESQL CONNECTION
# ============================================

engine = create_engine(
    "postgresql+psycopg2://postgres:3840@127.0.0.1:5432/arey_db"
)

# ============================================
# UPLOAD TO POSTGRESQL
# ============================================

df.to_sql(
    "personnel_movements",
    engine,
    if_exists="append",
    index=False
)

# ============================================
# SUCCESS MESSAGE
# ============================================

print("Дані успішно завантажені в PostgreSQL!")