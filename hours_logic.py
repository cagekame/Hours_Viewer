"""
hours_logic.py
==============
Logica applicativa: caricamento dati, settings, check ore.
Nessun import da tkinter.
"""

import json
import os
import subprocess
import gc
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

# ── Percorsi ──────────────────────────────────────────────────────────────────
_XLSX_DEFAULT = r"\\flowserve.net\des\CAD\lav\00_UTENTI\Caggianelli\Programmi\hours_data.xlsx"

def get_xlsx_path() -> str:
    """Legge il percorso del file Excel dal settings.json, altrimenti usa il default."""
    try:
        settings_file = Path(os.environ["USERPROFILE"]) / "Hours_Viewer" / "settings.json"
        with open(settings_file, "r", encoding="utf-8") as f:
            s = json.load(f)
        return s.get("xlsx_path", _XLSX_DEFAULT)
    except Exception:
        return _XLSX_DEFAULT

XLSX_FILE = get_xlsx_path()
# Percorso fisso - usa winreg per leggere il vero LOCALAPPDATA
# funziona sia da script, sia da exe PyInstaller, sia con Python Store
LOCAL_DIR     = Path(os.environ["USERPROFILE"]) / "Hours_Viewer"
JSON_FILE     = LOCAL_DIR / "hours_data.json"
SETTINGS_FILE = LOCAL_DIR / "settings.json"

DISPLAY_COLS = {
    "DescrResource": "Resource",
    "CDCOM":         "Job",
    "DescrFase":     "Phase",
    "DescrShort":    "Description",
    "WorkNote":      "Notes",
    "WorkDte":       "Date",
    "CreatedOn":     "Created On",
    "WorkHrs":       "Hours",
}

# ── Settings ──────────────────────────────────────────────────────────────────

def load_settings() -> dict:
    try:
        LOCAL_DIR.mkdir(parents=True, exist_ok=True)
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(settings: dict):
    try:
        LOCAL_DIR.mkdir(parents=True, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[WARN] Could not save settings: {e}")


# ── JSON helpers ──────────────────────────────────────────────────────────────

def json_to_df(json_path: Path = JSON_FILE):
    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    df = pd.DataFrame(payload["data"])
    df["WorkDte"]   = pd.to_datetime(df["WorkDte"],   errors="coerce").dt.date
    df["CreatedOn"] = pd.to_datetime(df["CreatedOn"], errors="coerce").dt.date
    df["WorkHrs"]   = pd.to_numeric(df["WorkHrs"],    errors="coerce").fillna(0.0)
    return df, payload.get("exported_at", "?")


def df_to_json(df: pd.DataFrame, json_path: Path = JSON_FILE):
    json_path.parent.mkdir(parents=True, exist_ok=True)
    rows = []
    for _, r in df.iterrows():
        row = {}
        for col in df.columns:
            val = r[col]
            if isinstance(val, (date, datetime)):
                val = str(val)
            elif pd.isna(val):
                val = None
            elif not isinstance(val, (int, float, bool, str)):
                val = str(val)
            row[col] = val
        rows.append(row)
    payload = {
        "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "count":       len(rows),
        "data":        rows,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# ── Excel refresh ─────────────────────────────────────────────────────────────

def full_refresh(xlsx_path: str, status_cb=None) -> pd.DataFrame:
    """Apre Excel in background, fa il refresh, legge i dati, chiude Excel."""
    import win32com.client
    import pythoncom
    import pywintypes

    pythoncom.CoInitialize()
    excel = None
    wb    = None
    excel_pid = None

    try:
        if status_cb: status_cb("Starting Excel in background...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible        = False
        excel.AskToUpdateLinks = False

        try:
            import win32process
            _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        except Exception:
            pass

        if status_cb: status_cb("Opening file...")
        wb = excel.Workbooks.Open(
            str(Path(xlsx_path).resolve()),
            UpdateLinks=False,
            ReadOnly=False
        )

        if status_cb: status_cb("Downloading data from database (~25 sec)...")
        for i in range(1, wb.Connections.Count + 1):
            c = wb.Connections(i)
            try:
                c.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass
            try:
                c.Refresh()
            except Exception:
                pass
        excel.CalculateUntilAsyncQueriesDone()

        if status_cb: status_cb("Reading data...")
        ws       = wb.Sheets(1)
        last_row = ws.UsedRange.Rows.Count
        last_col = ws.UsedRange.Columns.Count
        data     = ws.Range(ws.Cells(1,1), ws.Cells(last_row, last_col)).Value

        headers = list(data[0])

        def conv(val):
            if isinstance(val, pywintypes.TimeType):
                return f"{val.year}-{val.month:02d}-{val.day:02d}"
            return val

        rows = [[conv(v) for v in r] for r in data[1:]]
        return pd.DataFrame(rows, columns=headers)

    finally:
        try:
            if wb:    wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel: excel.Quit()
        except Exception:
            pass
        try:
            del wb, excel
            gc.collect()
        except Exception:
            pass
        pythoncom.CoUninitialize()
        if excel_pid:
            try:
                subprocess.run(["taskkill", "/F", "/PID", str(excel_pid)],
                               capture_output=True)
            except Exception:
                pass


# ── Hours check ───────────────────────────────────────────────────────────────

IT_HOLIDAYS_FIXED = {
    "01-01", "01-06", "04-25", "05-01", "06-02",
    "08-15", "11-01", "12-08", "12-25", "12-26"
}
HOURS_EXPECTED = {0: 8, 1: 8, 2: 8, 3: 8, 4: 6, 5: 0, 6: 0}
DAY_NAMES      = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def _easter(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day   = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def is_holiday(d: date) -> bool:
    if d.weekday() >= 5:
        return True
    if f"{d.month:02d}-{d.day:02d}" in IT_HOLIDAYS_FIXED:
        return True
    easter = _easter(d.year)
    if d == easter or d == easter + timedelta(days=1):
        return True
    return False


def check_hours(df: pd.DataFrame, risorsa: str, d_from: date, d_to: date) -> list:
    if not risorsa or risorsa == "All":
        return []
    df_r = df[
        (df["DescrResource"] == risorsa) &
        (df["WorkDte"] >= d_from) &
        (df["WorkDte"] <= d_to)
    ]
    ore_per_giorno = df_r.groupby("WorkDte")["WorkHrs"].sum().to_dict()
    issues = []
    current = d_from
    while current <= d_to:
        if not is_holiday(current):
            attese   = HOURS_EXPECTED.get(current.weekday(), 8)
            inserite = ore_per_giorno.get(current, 0)
            if inserite < attese:
                issues.append({
                    "data":         current,
                    "giorno":       DAY_NAMES[current.weekday()],
                    "ore_attese":   attese,
                    "ore_inserite": inserite,
                    "delta":        attese - inserite,
                })
        current += timedelta(days=1)
    return issues


def parse_date(s: str):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None