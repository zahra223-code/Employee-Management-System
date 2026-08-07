# Employee Dashboard, General Bureau of the East Java Provincial Secretariat

An internal Streamlit-based web application for managing employee data, monitoring employee performance (Working Hours KPI and 7-Aspect ASN/BerAKHLAK KPI), and automatically generating employee reports and report cards. This application was developed to replace the manual, spreadsheet-based recapitulation process that was previously scattered across multiple separate files, by consolidating the entire workflow — from employee data lookup, working hours discipline monitoring, seven-aspect ASN performance assessment, to retired employee archiving — into a single centralized dashboard.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Key Features](#key-features)
- [Architecture and Data Flow](#architecture-and-data-flow)
- [Project Directory Structure](#project-directory-structure)
- [Technology Stack](#technology-stack)
- [Installation and Running the Application](#installation-and-running-the-application)
- [Data Format Specifications](#data-format-specifications)
- [User Authentication](#user-authentication)
- [Application Menu Map](#application-menu-map)
- [Technical Notes and Design Decisions](#technical-notes-and-design-decisions)
- [Known Limitations and Risks](#known-limitations-and-risks)
- [Development Roadmap](#development-roadmap)
- [License and Contact](#license-and-contact)

---

## Project Overview

The Employee Dashboard is a single-page application built with Streamlit, functioning as an internal-scale employee information system without relying on relational database infrastructure. All data read and write operations are handled through CSV/XLSX files managed via `pandas`, allowing the application to run in resource-constrained computing environments without requiring a separate database server installation. This is a deliberate trade-off: deployment simplicity and portability are prioritized over simultaneous multi-user scalability — this consideration is discussed further in the [Known Limitations and Risks](#known-limitations-and-risks) section.

---

## Key Features

### Employee Data Management

- Real-time employee search by name, employee ID (NIP), or national ID (NIK).
- Addition, modification, and deletion (archiving) of employee data through a form-based interface, eliminating the need for manual editing of CSV/Excel files.
- Automatic retirement detection and archiving: employees who have passed their retirement effective date (TMT) are automatically moved to the Employee Archive each time the application loads, ensuring active data consistently reflects actual employment status.
- Automatic normalization of column naming variations in master files (e.g., `GOL. RUANG` vs. `GOLONGAN`), maintaining consistency when data is imported from different sources.

### Employee Management

- Structured input form for adding new employee records, with required-field validation before data is saved to `pegawai_kelola.csv`.
- Employee data update mechanism (biodata, position, work unit, rank, and other employment attributes) without altering other employees' rows.
- Manual archiving function, allowing users to move an employee to the Employee Archive directly (outside the automatic retirement detection mechanism), for cases such as transfer, resignation, or dismissal.
- Every data change made through this menu is recorded in the User Activity Log, ensuring employee data changes remain traceable.

### Employee Information Dashboard

- Interactive visualizations using Plotly for age composition, gender, education level, religion, employment status, as well as work unit and position distribution.
- Quantitative summary of active employee counts, updated in real time based on all filters applied by the user.

### Working Hours KPI

- Upload raw (daily) attendance data in CSV or XLSX format, automatically processed into a discipline score per employee.
- Automatic calculation for nine violation components: Permission Leave, Late Arrival, Early Departure, No Clock-In, No Clock-Out, Absent Without Notice, Missed Exercise Session, Late for Exercise Session, and Missed Roll Call, each with its own deduction rate.
- Flexible period filtering: all data, a specific month, or a custom date range.
- Performance status distribution chart, a leaderboard of the fifteen employees with the highest deductions, and per-employee data breakdown.
- Automatic filtering based on employment status, ensuring retired, deleted, or archived employees do not affect the statistics.
- Export of recapitulation to CSV format.

### 7-Aspect ASN KPI (BerAKHLAK)

- Upload of seven-aspect assessments (Service Orientation, Accountable, Competent, Harmonious, Loyal, Adaptive, Collaborative) per period, with an upsert mechanism: data with an identical combination of NIP, Name, and Period will be updated rather than duplicated.
- KPI summary cards (Achieved, Below Expectation, Average Score) calculated from a single most recent data row per employee, avoiding calculation bias when an employee has data from multiple periods simultaneously.
- Radar chart of average scores per aspect and pie chart of status distribution, numerically consistent with the figures shown on the summary cards.
- Detailed per-employee view in "Monthly" or "Overall Average" mode.
- Automatic filtering based on employment status, accompanied by explicit warning messages when uploaded data does not match employee records (e.g., due to name spelling errors).

### Employee Report

- Combines biodata, Working Hours KPI summary, and 7-Aspect KPI summary into a single unified view per employee.
- Export of individual reports to print-ready PDF format, built using ReportLab.

### Employee Archive

- Stores the history of employees who have retired, transferred, resigned, or been deactivated, kept separate from active data but remaining traceable.
- Export of archive data to CSV format.

### User Activity Log

- Records user activity trails during an active session (login, data upload, report download, employee data changes, logout, etc.).
- Export of the activity log to CSV format.

### Authentication and Interface Consistency

- Session-based login system (`st.session_state`) with a dedicated login page.
- Interface design consistent with the government institution's visual identity (blue color palette, Poppins typography), including a navigation sidebar optimized to maintain text readability across all input component types.

---

## Architecture and Data Flow

This application does not use a relational database management system. All operational data is stored as CSV files within a local directory structure, and is read from and written back to by the application via `pandas`. This architectural decision is based on the requirement that the application run internally without needing a separate database server installation, with the logical consequence of limitations in concurrent access scenarios (see [Known Limitations and Risks](#known-limitations-and-risks)).

The general data flow can be illustrated as follows:
```
Data_Master.xlsx (data awal)
        |
        v
pegawai_kelola.csv  <──────────────┐
        |                          │  (changes saved automatically)
        v                          │
  [ Aplikasi Streamlit ] ──────────┘
        |
        ├── kehadiran_final.csv      → processed into Working Hours KPI
        ├── kpi_7aspek_final.csv     → processed into 7-Aspect ASN KPI
        └── arsip_pegawai.csv        ← retired/inactive employees moved automatically
```

Every time the application loads, active employee data is re-filtered against retirement effective dates (TMT), and employees who have passed their term of duty are automatically moved to the archive. This mechanism ensures that all statistical calculations and KPIs on the dashboard consistently represent the genuinely active employee population, rather than uncleaned historical data.

---

## Project Directory Structure

The application assumes the following directory structure relative to the project root (see the `BASE_DIR` configuration in the source code):
---


```
project-root/
├── 01_source_code/
│   └── dashboard_kepegawaian.py        # Application entry point
├── 03_dataset/
│   └── employee/
│       └── Data_Master.xlsx            # Initial employee data (seed)
├── 04_database_operasional/
│   ├── employee/
│   │   ├── pegawai_kelola.csv          # Active (live) employee data
│   │   └── arsip_pegawai.csv           # Inactive employee archive
│   ├── attendance/
│   │   └── kehadiran_final.csv         # Daily attendance database
│   ├── performance/
│   │   └── kpi_7aspek_final.csv        # 7-aspect ASN assessment database
│   └── activity/
│       └── log_aktivitas.csv           # No longer used; log is now in-memory
└── 05_asset/
    └── LOGO.png                        # Institution logo for login page/sidebar
```

The `04_database_operasional/` folder contains sensitive employee data and must not be uploaded to a public repository. Use a `.gitignore` file to explicitly exclude this folder (see the [Installation](#installation-and-running-the-application) section).

---

## Technology Stack

| Functional Requirement | Library/Framework |
|---|---|
| Web framework and dashboard | [Streamlit](https://streamlit.io/) |
| Tabular data manipulation | [pandas](https://pandas.pydata.org/), [numpy](https://numpy.org/) |
| Interactive visualization | [Plotly](https://plotly.com/python/) (`plotly.express`, `plotly.graph_objects`) |
| Excel file reading and writing | [openpyxl](https://openpyxl.readthedocs.io/) (engine dependency for `pandas`) |
| PDF document generation | [ReportLab](https://www.reportlab.com/) |
| Precise date calculations | [python-dateutil](https://dateutil.readthedocs.io/) |
| Cross-platform path management | `pathlib` (Python built-in module) |

---

## Installation and Running the Application

### 1. System Requirements

- Python version 3.9 or later.
- `pip` installed on the system.

### 2. Clone the Repository

```bash
git clone https://github.com/<username>/<repo-name>.git
cd <repo-name>
```

### 3. Create a Virtual Environment

Using a virtual environment is strongly recommended to isolate project dependencies from the system Python installation.

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 4. Install Dependencies

Create a `requirements.txt` file (if not already available) with the following content:

```
streamlit
pandas
numpy
plotly
reportlab
python-dateutil
openpyxl
```

Then run:

```bash
pip install -r requirements.txt
```

### 5. Prepare the Data Directory Structure

Ensure the directory structure matches the description in the [Project Directory Structure](#project-directory-structure) section, with at minimum the following two files available:

- `03_dataset/employee/Data_Master.xlsx`, the initial employee data.
- `05_asset/LOGO.png`, the institution logo.

Files within `04_database_operasional/` will be automatically created and populated by the application as it is used (data uploads, additions/changes, etc.).

### 6. Run the Application

```bash
streamlit run 01_source_code/dashboard_kepegawaian.py
```

The application will open automatically at `http://localhost:8501`.

---

## Data Format Specifications

### Employee Master Data (`Data_Master.xlsx`)

Minimum expected columns (column naming variations are automatically normalized by the application):

`NO`, `NAMA`, `NIP BARU`, `NIK`, `JK`, `AGAMA`, `STATUS PNS/CPNS`, `JENJANG`, `NAMA JABATAN`, `UNIT KERJA`, `TGL LAHIR`, `USIA`, `TMT MASUK KERJA`, `TMT PANGKAT`, `GOLONGAN`, `TMT PENSIUN`, `TAHUN PENSIUN`.

### Attendance Data (uploaded via the Working Hours KPI menu)

CSV or XLSX format containing daily attendance data per employee, including columns such as `nama`, `nip`, `opd`, `tanggal`, `kode_asli`, `menit_terlambat`, and `potong_gaji`. Attendance codes are automatically mapped to the nine violation components (Permission Leave, Late Arrival, and so on).

### 7-Aspect ASN Assessment Data (uploaded via the 7-Aspect ASN KPI menu)

Required columns: `NAMA`, `Orientasi_Pelayanan`, `Akuntabel`, `Kompeten`, `Harmonis`, `Loyal`, `Adaptif`, `Kolaboratif`.

Optional columns: `NIP`, `Periode` (format `YYYY-MM`, e.g., `2026-01`).

---

## User Authentication

The application uses a Python dictionary-based login system (`USERS`) validated through `st.session_state`, suitable for a limited internal user base with minimal security requirements.

**Important note for production environments.** The current implementation stores credentials as plain text within the source code. This approach is adequate for the development or demonstration stage, but is not suitable for production deployment without modification. Recommended improvements include:

- Moving credentials to environment variables or a secrets manager (e.g., `st.secrets` on Streamlit Cloud).
- Implementing password hashing (e.g., using `bcrypt` or `passlib`).
- Ensuring source files containing actual credentials are not included in a public repository.

---

## Application Menu Map

| Menu | Function |
|---|---|
| Employee Information | Employee data search and detail view |
| Employee Information Dashboard | Visualization of employee demographic and organizational composition |
| Employee Management | Input, update, and deletion (archiving) of employee data |
| Working Hours KPI | Attendance upload, discipline score calculation, and recapitulation |
| 7-Aspect ASN KPI | BerAKHLAK 7-aspect assessment upload and performance recapitulation |
| Employee Report | Per-employee biodata and KPI summary, PDF export |
| Employee Archive | History of inactive, retired, or transferred employees |
| Activity Log | User activity log for the active session |
| Logout | Session exit confirmation and logout |

---

## Technical Notes and Design Decisions

- **KPI statistical consistency.** All summary cards, pie charts, and radar charts on the KPI menus are calculated from a single, identical source row of data per employee, rather than from all raw rows, preventing double-counting bias when an employee has more than one period of data in "All Data" mode.
- **Employment status synchronization.** KPI data, both Working Hours and 7-Aspect, is consistently filtered against the list of genuinely active employees in the employee database, not merely based on uploaded data. This mechanism prevents retired or deactivated employees from affecting performance statistics.
- **Error message differentiation.** The application explicitly distinguishes between the condition "database has never been populated" and "data is available but has been entirely filtered out due to a mismatch with active employee records" (e.g., due to name spelling errors), enabling faster and more accurate debugging by users and IT teams alike.
- **Style (CSS) isolation.** Sidebar and main content area styling rules are explicitly separated to prevent newly added input components from inheriting styles that are inappropriate for their context (e.g., dark text on a dark background).

---

## Known Limitations and Risks

This section is included deliberately, so that users and future developers understand the architectural trade-offs inherent in the application's current design, rather than simply reading a list of shortcomings.

- **Access concurrency.** Because storage is CSV-file-based without an explicit locking mechanism, simultaneous data writes by more than one user could potentially cause race conditions or data loss. This application is best suited for sequential usage scenarios or a small number of concurrently active users.
- **Data volume scalability.** The approach of loading entire CSV files into memory (`pandas.read_csv`) will experience performance degradation as employee data volume grows over the long term.
- **Credential security.** As explained in the [User Authentication](#user-authentication) section, the current authentication scheme does not yet meet production security standards.

---

## Development Roadmap

- Migrate data storage from CSV to a database system (PostgreSQL or SQLite) to support concurrent multi-user access.
- Implement user credential hashing and multi-role support (admin, staff, management).
- Export Working Hours and 7-Aspect KPI recapitulations to Excel (`.xlsx`) format, complementing the existing CSV format.
- Automatic in-app notifications for employees approaching retirement within an upcoming period.
- Development of automated tests (unit tests) for KPI data processing functions.

---

## License and Contact

This project was developed for the internal needs of the General Bureau of the Regional Secretariat. The licensing terms in this section should be adjusted according to the relevant institution's or organization's policy (e.g., proprietary/internal use only, or an open-source license such as MIT if the project is to be shared publicly).

For questions, bug reports, or feature requests, please open an Issue on this repository or contact the internal development team.
Untuk pertanyaan, laporan bug, atau permintaan fitur, silakan membuka Issue pada repositori ini atau menghubungi tim pengembang internal.
