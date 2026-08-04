# ==========================================================
# IMPORT LIBRARY
# ==========================================================

from pathlib import Path
from typing import Dict, List

import re

import numpy as np
import pandas as pd


# ==========================================================
# CONFIGURATION
# ==========================================================

"""
Konfigurasi lokasi folder project.
"""

PROJECT_FOLDER = Path(__file__).resolve().parent.parent

DATASET_FOLDER = PROJECT_FOLDER / "03_dataset"

EMPLOYEE_FOLDER = DATASET_FOLDER / "employee"

MASTER_FILE = EMPLOYEE_FOLDER / "Data_Master.xlsx"

ATTENDANCE_INPUT_FOLDER = (
    DATASET_FOLDER
    / "attendance_mentah"
)

ATTENDANCE_OUTPUT_FOLDER = (
    DATASET_FOLDER
    / "attendance_anonymized"
)


# ==========================================================
# CONSTANTS
# ==========================================================

"""
Konfigurasi Data Master.
"""

MASTER_SHEET_NAME = "Data_Master"

MASTER_COLUMN_OLD_NIP = "NIP_Lama"

MASTER_COLUMN_NEW_NIP = "NIP"

MASTER_COLUMN_ALIAS_NAME = "Nama_Samaran"

MASTER_COLUMN_RETIREMENT_DATE = "TMT_Pensiun"


"""
Konfigurasi Attendance.
"""

ATTENDANCE_COLUMN_NIP = "NIP"

ATTENDANCE_COLUMN_NAME = "Nama Pegawai"


"""
Konfigurasi File.
"""

SUPPORTED_FILE_EXTENSION = [
    ".xlsx",
    ".xls",
    ".csv"
]

OUTPUT_ENGINE = "openpyxl"

SAVE_INDEX = False


# ==========================================================
# UTILITY FUNCTIONS
# ==========================================================

def print_header(title: str) -> None:
    """
    Menampilkan judul proses pada terminal.
    """

    print("\n" + "=" * 70)
    print(title.upper())
    print("=" * 70)


def ensure_output_folder() -> None:
    """
    Membuat folder output apabila belum tersedia.
    """

    ATTENDANCE_OUTPUT_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )


def clean_nip(value):
    """
    Membersihkan format NIP agar memiliki format yang seragam.
    """

    if pd.isna(value):
        return ""

    value = str(value)

    value = value.replace("`", "")
    value = value.replace("'", "")
    value = value.replace('"', "")

    value = value.strip()

    value = re.sub(
        r"\s+",
        "",
        value
    )

    return value


def normalize_column_name(column: str) -> str:
    """
    Menyeragamkan nama kolom.
    """

    column = str(column)

    column = column.strip()

    column = re.sub(
        r"\s+",
        " ",
        column
    )

    return column


def normalize_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    Menyeragamkan seluruh nama kolom dataframe.
    """

    dataframe.columns = [
        normalize_column_name(column)
        for column in dataframe.columns
    ]

    return dataframe


def load_master_data() -> pd.DataFrame:
    """
    Membaca Data Master.
    """

    print_header("Load Data Master")

    master_dataframe = pd.read_excel(
        MASTER_FILE,
        sheet_name=MASTER_SHEET_NAME,
        dtype=str
    )

    master_dataframe = normalize_dataframe(
        master_dataframe
    )

    master_dataframe[MASTER_COLUMN_OLD_NIP] = (
        master_dataframe[
            MASTER_COLUMN_OLD_NIP
        ].apply(clean_nip)
    )

    master_dataframe[MASTER_COLUMN_NEW_NIP] = (
        master_dataframe[
            MASTER_COLUMN_NEW_NIP
        ].apply(clean_nip)
    )

    print(
        f"Jumlah Pegawai : {len(master_dataframe)}"
    )

    return master_dataframe


def get_attendance_files() -> List[Path]:
    """
    Mengambil seluruh file attendance.
    """

    print_header("Scanning Attendance Folder")

    attendance_files = []

    for file in ATTENDANCE_INPUT_FOLDER.iterdir():

        if (
            file.is_file()
            and file.suffix.lower()
            in SUPPORTED_FILE_EXTENSION
        ):
            attendance_files.append(file)

    attendance_files.sort()

    print(
        f"Jumlah File : {len(attendance_files)}"
    )

    return attendance_files


def load_attendance_file(
    file_path: Path
) -> pd.DataFrame:
    """
    Membaca satu file attendance.
    """

    print(
        f"Membaca : {file_path.name}"
    )

    extension = file_path.suffix.lower()

    if extension == ".csv":

        attendance_dataframe = pd.read_csv(
            file_path,
            dtype=str
        )

    elif extension in [".xlsx", ".xls"]:

        attendance_dataframe = pd.read_excel(
            file_path,
            sheet_name=0,
            dtype=str
        )

    else:

        raise ValueError(
            f"Format file tidak didukung : {extension}"
        )

    attendance_dataframe = normalize_dataframe(
        attendance_dataframe
    )

    return attendance_dataframe

def save_attendance_file(
    dataframe: pd.DataFrame,
    output_file: Path
) -> None:
    """
    Menyimpan hasil anonymisasi.
    """

    dataframe.to_excel(
        output_file,
        engine=OUTPUT_ENGINE,
        index=SAVE_INDEX
    )

    print(
        f"Berhasil Disimpan : {output_file.name}"
    )


def print_summary(
    file_name: str,
    before_rows: int,
    after_rows: int
) -> None:
    """
    Menampilkan ringkasan proses.
    """

    print("-" * 70)

    print(
        f"File           : {file_name}"
    )

    print(
        f"Data Awal      : {before_rows}"
    )

    print(
        f"Data Akhir     : {after_rows}"
    )

    print("-" * 70)

# ==========================================================
# VALIDATION ENGINE
# ==========================================================

def validate_master_dataframe(
    master_dataframe: pd.DataFrame
) -> None:
    """
    Memastikan struktur Data Master sesuai.
    """

    required_columns = [
        MASTER_COLUMN_OLD_NIP,
        MASTER_COLUMN_NEW_NIP,
        MASTER_COLUMN_ALIAS_NAME,
        MASTER_COLUMN_RETIREMENT_DATE
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in master_dataframe.columns
    ]

    if missing_columns:

        raise ValueError(
            f"Kolom Data Master tidak ditemukan: {missing_columns}"
        )


def validate_attendance_dataframe(
    attendance_dataframe: pd.DataFrame
) -> None:
    """
    Memastikan struktur Attendance sesuai.
    """

    required_columns = [
        ATTENDANCE_COLUMN_NIP,
        ATTENDANCE_COLUMN_NAME
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in attendance_dataframe.columns
    ]

    if missing_columns:

        raise ValueError(
            f"Kolom Attendance tidak ditemukan: {missing_columns}"
        )


def validate_empty_master(
    master_dataframe: pd.DataFrame
) -> None:
    """
    Memastikan Data Master tidak kosong.
    """

    if master_dataframe.empty:

        raise ValueError(
            "Data Master kosong."
        )


def validate_empty_attendance(
    attendance_dataframe: pd.DataFrame
) -> None:
    """
    Memastikan Attendance tidak kosong.
    """

    if attendance_dataframe.empty:

        raise ValueError(
            "Attendance kosong."
        )


def validate_duplicate_master(
    master_dataframe: pd.DataFrame
) -> None:
    """
    Memastikan NIP Lama tidak memiliki data duplikat.
    """

    duplicate_count = master_dataframe[
        MASTER_COLUMN_OLD_NIP
    ].duplicated().sum()

    if duplicate_count > 0:

        raise ValueError(
            f"Ditemukan {duplicate_count} NIP Lama duplikat."
        )


# ==========================================================
# LOOKUP ENGINE
# ==========================================================

def create_lookup_table(
    master_dataframe: pd.DataFrame
) -> Dict[str, dict]:
    """
    Membuat lookup table berdasarkan NIP Lama.
    """

    lookup_table = {}

    for _, row in master_dataframe.iterrows():

        lookup_table[
            row[MASTER_COLUMN_OLD_NIP]
        ] = {

            "new_nip":
            row[MASTER_COLUMN_NEW_NIP],

            "alias_name":
            row[MASTER_COLUMN_ALIAS_NAME],

            "retirement_date":
            row[MASTER_COLUMN_RETIREMENT_DATE]

        }

    return lookup_table


def clean_attendance_nip(
    attendance_dataframe: pd.DataFrame
) -> pd.DataFrame:
    """
    Membersihkan seluruh nilai NIP pada Attendance.
    """

    attendance_dataframe[
        ATTENDANCE_COLUMN_NIP
    ] = attendance_dataframe[
        ATTENDANCE_COLUMN_NIP
    ].apply(clean_nip)

    return attendance_dataframe


def get_lookup_record(
    lookup_table: Dict[str, dict],
    old_nip: str
):
    """
    Mengambil data pegawai dari lookup table.
    """

    return lookup_table.get(
        old_nip,
        None
    )


def count_unregistered_employee(
    attendance_dataframe: pd.DataFrame,
    lookup_table: Dict[str, dict]
) -> int:
    """
    Menghitung jumlah NIP Attendance yang tidak ditemukan
    pada Data Master.
    """

    total = 0

    for nip in attendance_dataframe[
        ATTENDANCE_COLUMN_NIP
    ]:

        if nip not in lookup_table:

            total += 1

    return total

# ==========================================================
# ANONYMIZATION ENGINE
# ==========================================================

def anonymize_attendance(
    attendance_dataframe: pd.DataFrame,
    lookup_table: Dict[str, dict]
):
    """
    Melakukan proses anonymisasi dataset attendance.
    """

    anonymized_rows = []

    total_rows = len(attendance_dataframe)

    mapped_rows = 0

    removed_unregistered = 0

    removed_retired = 0

    today = pd.Timestamp.today().normalize()

    for _, row in attendance_dataframe.iterrows():

        old_nip = row[ATTENDANCE_COLUMN_NIP]

        employee = get_lookup_record(
            lookup_table,
            old_nip
        )

        if employee is None:

            removed_unregistered += 1

            continue

        retirement_date = employee[
            "retirement_date"
        ]

        if pd.notna(retirement_date):

            retirement_date = pd.to_datetime(
                retirement_date,
                errors="coerce"
            )

            if retirement_date <= today:

                removed_retired += 1

                continue

        row[
            ATTENDANCE_COLUMN_NIP
        ] = employee[
            "new_nip"
        ]

        row[
            ATTENDANCE_COLUMN_NAME
        ] = employee[
            "alias_name"
        ]

        anonymized_rows.append(
            row
        )

        mapped_rows += 1

    anonymized_dataframe = pd.DataFrame(
        anonymized_rows
    )

    statistics = {

        "total_rows": total_rows,

        "mapped_rows": mapped_rows,

        "removed_unregistered": removed_unregistered,

        "removed_retired": removed_retired,

        "final_rows": len(
            anonymized_dataframe
        )

    }

    return (
        anonymized_dataframe,
        statistics
    )


def print_anonymization_summary(
    file_name: str,
    statistics: dict
) -> None:
    """
    Menampilkan ringkasan hasil anonymisasi.
    """

    print("\n" + "=" * 70)

    print(
        f"FILE                 : {file_name}"
    )

    print(
        f"DATA AWAL            : {statistics['total_rows']}"
    )

    print(
        f"BERHASIL DIMAPPING   : {statistics['mapped_rows']}"
    )

    print(
        f"TIDAK TERDAFTAR      : {statistics['removed_unregistered']}"
    )

    print(
        f"SUDAH PENSIUN        : {statistics['removed_retired']}"
    )

    print(
        f"DATA AKHIR           : {statistics['final_rows']}"
    )

    print("=" * 70)

# ==========================================================
# SAVE ENGINE
# ==========================================================

def generate_output_path(
    input_file: Path
) -> Path:
    """
    Membuat lokasi penyimpanan file hasil anonymisasi.
    """

    return ATTENDANCE_OUTPUT_FOLDER / input_file.name


def save_anonymized_attendance(
    anonymized_dataframe: pd.DataFrame,
    input_file: Path,
    statistics: dict
) -> None:
    """
    Menyimpan hasil anonymisasi ke folder output.
    """

    ensure_output_folder()

    output_file = generate_output_path(
        input_file
    )

    save_attendance_file(
        anonymized_dataframe,
        output_file
    )

    print_anonymization_summary(
        input_file.name,
        statistics
    )


def process_single_attendance(
    file_path: Path,
    lookup_table: Dict[str, dict]
) -> None:
    """
    Memproses satu file attendance.
    """

    attendance_dataframe = load_attendance_file(
        file_path
    )

    validate_attendance_dataframe(
        attendance_dataframe
    )

    validate_empty_attendance(
        attendance_dataframe
    )

    attendance_dataframe = clean_attendance_nip(
        attendance_dataframe
    )

    anonymized_dataframe, statistics = anonymize_attendance(
        attendance_dataframe,
        lookup_table
    )

    save_anonymized_attendance(
        anonymized_dataframe,
        file_path,
        statistics
    )


def process_all_attendance(
    lookup_table: Dict[str, dict]
) -> None:
    """
    Memproses seluruh file attendance.
    """

    attendance_files = get_attendance_files()

    if len(attendance_files) == 0:

        print(
            "\nTidak ada file attendance."
        )

        return

    print_header(
        "Process Attendance Dataset"
    )

    for file_path in attendance_files:

        process_single_attendance(
            file_path,
            lookup_table
        )

    print_header(
        "Process Finished"
    )

    print(
        "Seluruh file berhasil diproses."
    )

# ==========================================================
# MAIN PIPELINE
# ==========================================================

def main() -> None:
    """
    Menjalankan seluruh proses anonymisasi attendance.
    """

    print_header(
        "Attendance Anonymization Tool"
    )

    ensure_output_folder()

    master_dataframe = load_master_data()

    validate_master_dataframe(
        master_dataframe
    )

    validate_empty_master(
        master_dataframe
    )

    validate_duplicate_master(
        master_dataframe
    )

    lookup_table = create_lookup_table(
        master_dataframe
    )

    process_all_attendance(
        lookup_table
    )

    print_header(
        "Process Completed"
    )

    print(
        "Seluruh proses anonymisasi berhasil diselesaikan."
    )


# ==========================================================
# ENTRY POINT
# ==========================================================

if __name__ == "__main__":

    main()