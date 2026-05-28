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
# READ EXCEL
# ============================================

print("\n===================================")
print("ЧИТАННЯ EXCEL ФАЙЛУ")
print("===================================\n")

df = pd.read_excel(
    file_path,
    sheet_name="ООС",
    header=2,
    engine="openpyxl",
    dtype={
        "20": str,
        "302": str,
        "308": str,
        "21": str
    }
)

print(f"Зчитано рядків з Excel: {len(df)}")
print(f"Зчитано колонок: {len(df.columns)}")

# ============================================
# CONVERT COLUMN NAMES TO STRING
# ============================================

df.columns = df.columns.map(str)

# ============================================
# COLUMN MAPPING
# ============================================

column_mapping = {

    "9": "full_name",
    "8": "rank_name",
    "301": "callsign",
    "20": "tax_number",
    "302": "phone",
    "303": "unit_status",
    "304": "status",
    "400": "shtat_position_number",
    "305": "current_position_index",
    "10": "position_index_history",
    "11": "appointment_date",
    "11.1": "appointment_order_number",
    "12": "arrived_from_unit",
    "13": "enlistment_date",
    "14": "enlistment_order_date",
    "14.0": "enlistment_order_number",
    "15": "appointment_order_date",
    "15.0": "appointment_order_full",
    "15.1": "appointment_order_incoming_date",
    "15.2": "appointment_order_incoming_number",
    "16": "last_rank_order_date",
    "16.0": "last_rank_order_full",
    "16.1": "last_rank_order_incoming_date",
    "16.2": "last_rank_order_incoming_number",
    "17": "military_service_type",
    "18": "contract_start_date",
    "19": "contract_end_date",
    "306": "note",
    "22": "birth_date",
    "307": "age",
    "401": "citizenship",
    "23": "birth_place",
    "24": "gender",
    "21": "identity_document_number",
    "21.0": "identity_document_type",
    "308": "military_document_number",
    "25": "recruiting_office",
    "25.0": "mobilization_date",
    "26": "education",
    "27": "family_info",
    "309": "marital_status",
    "310": "clothing_size",
    "311": "blood_group",
    "312": "token_status",
    "313": "registration_address",
    "314": "residence_address",
    "315": "relatives_in_occupied_territory",
    "316": "service_period",
    "317": "ato_oos_participation",
    "318": "ubd_status",
    "318_1": "ubd_issue_date",
    "318_2": "ubd_order_number",
    "319": "weapon_info",
    "320": "bzvp_info",
    "358": "certificates_info",
    "321": "internally_displaced_person",
    "322": "disability_info",
    "323": "total_service_years",
    "324": "parent_unit_note",
    "357": "vlk_certificate",
    "325": "fitness_for_service",
    "28": "additional_info",
    "356": "army_plus_id",
    "328": "exclusion_date",
    "329": "exclusion_order_date",
    "330": "exclusion_order_number",
    "338": "exclusion_reason",
    "331": "loss_type",
    "332": "loss_date",
    "333": "loss_circumstances",
    "334": "loss_place",
    "335": "notification_info",
    "336": "burial_info",
    "337": "documents_sent_info",
    "350": "disposition_date",
    "351": "disposition_order_number",
    "353": "disposition_reason",
    "354": "personal_documents_status",
    "406": "surname",
    "407": "initials",
    "304_0": "concentration_area_status",
    "408": "rank_category_short",
    "402": "old_value",
    "403": "new_value",
    "404": "difference_flag",
    "405": "merged_value"
}

# ============================================
# KEEP ONLY NEEDED COLUMNS
# ============================================

print("\n===================================")
print("ФІЛЬТРАЦІЯ КОЛОНОК")
print("===================================\n")

df = df[list(column_mapping.keys())]

print(f"Залишено колонок: {len(df.columns)}")

# ============================================
# RENAME COLUMNS
# ============================================

df = df.rename(columns=column_mapping)

print("\nКолонки успішно перейменовані")

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
    subset=["full_name"]
)

# ============================================
# REMOVE SERVICE ROWS
# ============================================

df = df[
    df["full_name"].astype(str).str.strip() != ""
]

# ============================================
# TEXT COLUMNS
# ============================================

text_columns = [

    col for col in df.columns

    if col not in [

        "appointment_date",
        "enlistment_date",
        "enlistment_order_date",
        "appointment_order_date",
        "appointment_order_incoming_date",
        "last_rank_order_date",
        "last_rank_order_incoming_date",
        "contract_start_date",
        "birth_date",
        "mobilization_date",
        "ubd_issue_date",
        "exclusion_date",
        "exclusion_order_date",
        "loss_date",
        "disposition_date",
        "difference_flag",
        "age"
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

    "appointment_date",
    "enlistment_date",
    "enlistment_order_date",
    "appointment_order_date",
    "appointment_order_incoming_date",
    "last_rank_order_date",
    "last_rank_order_incoming_date",
    "contract_start_date",
    "birth_date",
    "mobilization_date",
    "ubd_issue_date",
    "exclusion_date",
    "exclusion_order_date",
    "loss_date",
    "disposition_date"
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
# AGE
# ============================================

df["age"] = pd.to_numeric(
    df["age"],
    errors="coerce"
)

# ============================================
# BOOLEAN
# ============================================

print("\n===================================")
print("ОБРОБКА BOOLEAN")
print("===================================\n")

df["difference_flag"] = (

    df["difference_flag"]

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
# DEBUG
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
                "SELECT COUNT(*) FROM personnel_main"
            )
        ).scalar()

        # ====================================
        # CLEAR TABLE
        # ====================================

        conn.execute(
            text(
                "TRUNCATE TABLE personnel_main"
            )
        )

        print("Таблиця personnel_main очищена")

        # ====================================
        # UPLOAD DATA
        # ====================================

        print("\nПочаток завантаження даних...\n")

        df.to_sql(
            "personnel_main",
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

    print(f"Таблиця PostgreSQL: personnel_main")

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