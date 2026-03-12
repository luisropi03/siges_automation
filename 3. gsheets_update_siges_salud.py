# gsheets_update_siges_salud.py
# pip install gspread google-auth pandas openpyxl

import os
import re
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ====== Paths y configuración ======
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "reporte_siges_salud.xlsx"
EXCEL_SHEET = "detallexproyecto"
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "pmo-471920-e3d1ef30d163.json")

SPREADSHEET_ID = "1fA9Aj7Z-Nu96LBghzD5-czdaWWcNcQ-sJG3LNrnNv5s"
SHEET_NAME_DEST = "siges_salud"

# ====== Utilidades ======
ZWSP = re.compile(r"[\u200B\u200C\u200D\uFEFF\u00A0]")

def norm_str(s):
    if pd.isna(s):
        return ""
    s = str(s)
    s = ZWSP.sub("", s).strip()
    return s

def _excel_serial_to_dt(n: float):
    bases = [datetime(1899, 12, 30), datetime(1904, 1, 1)]
    for base in bases:
        dt = base + timedelta(days=float(n))
        if 1990 <= dt.year <= 2100:
            return dt
    return bases[0] + timedelta(days=float(n))

def to_periodo(val) -> str:
    if pd.isna(val):
        return ""

    if isinstance(val, (pd.Timestamp, datetime)):
        return pd.to_datetime(val).strftime("%Y-%m")

    s = str(val).strip()

    m = re.match(r"^\s*(\d{4})[-/](\d{1,2})\s*$", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"

    m = re.match(r"^\s*(\d{4})(\d{2})\s*$", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"

    m = re.match(r"^\s*(\d{1,2})[-/](\d{4})\s*$", s)
    if m:
        return f"{m.group(2)}-{int(m.group(1)):02d}"

    try:
        num = float(s)
        if num > 1e11:
            return pd.to_datetime(num, unit="ms", origin="unix").strftime("%Y-%m")
        if num > 1e9:
            return pd.to_datetime(num, unit="s", origin="unix").strftime("%Y-%m")
        if 0 < num < 100000:
            dt = _excel_serial_to_dt(num)
            return dt.strftime("%Y-%m")
    except Exception:
        pass

    try:
        return pd.to_datetime(s, dayfirst=True, errors="raise").strftime("%Y-%m")
    except Exception:
        return ""

def connect_gsheets(sa_file: str):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(sa_file, scopes=scopes)
    return gspread.authorize(creds)

def read_excel_detalle(path: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name)

    col_map = {
        "PRYcodigo": "codigo_proyecto_salud",
        "ProcesoRecurso": "codigo_proceso_recurso",
        "PRYTcodigo": "prot_codigo",
        "PRYRAfecha": "periodo",
        "PRYRAhoras": "horas_recibidas_salud",
    }
    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    required = list(col_map.values())
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise RuntimeError(f"Faltan columnas en el Excel origen: {missing}")

    df["codigo_proyecto_salud"]  = df["codigo_proyecto_salud"].apply(norm_str)
    df["codigo_proceso_recurso"] = df["codigo_proceso_recurso"].apply(norm_str)
    df["prot_codigo"]            = df["prot_codigo"].apply(norm_str)
    df["periodo"]                = df["periodo"].apply(to_periodo)
    df["horas_recibidas_salud"]  = pd.to_numeric(
        df["horas_recibidas_salud"], errors="coerce"
    ).fillna(0.0)

    df = df[
        (df["codigo_proyecto_salud"] != "")
        & (df["codigo_proceso_recurso"] != "")
        & (df["prot_codigo"] != "")
        & (df["periodo"] != "")
    ]

    return df[required].copy()

def read_ws_as_df(ws) -> pd.DataFrame:
    values = ws.get_all_values()
    header = [
        "codigo_proyecto_salud",
        "codigo_proceso_recurso",
        "prot_codigo",
        "periodo",
        "horas_recibidas_salud",
    ]
    if not values:
        return pd.DataFrame(columns=header)

    h, rows = values[0], values[1:]
    df = pd.DataFrame(rows, columns=h)

    for c in header[:-1]:
        df[c] = df.get(c, "").apply(norm_str)

    df["horas_recibidas_salud"] = pd.to_numeric(
        df.get("horas_recibidas_salud", 0),
        errors="coerce"
    ).fillna(0.0)

    return df[header].copy()

def write_df_to_ws(ws, df: pd.DataFrame):
    header = [
        "codigo_proyecto_salud",
        "codigo_proceso_recurso",
        "prot_codigo",
        "periodo",
        "horas_recibidas_salud",
    ]
    ws.clear()
    ws.update([header] + df.astype(object).values.tolist())

def upsert_destino(df_dest: pd.DataFrame, df_src: pd.DataFrame) -> pd.DataFrame:
    key_cols = [
        "codigo_proyecto_salud",
        "codigo_proceso_recurso",
        "prot_codigo",
        "periodo",
    ]

    dest_idx = df_dest.set_index(key_cols)
    src_idx = df_src.set_index(key_cols)

    dest_idx.update(src_idx)
    combined = pd.concat(
        [dest_idx, src_idx[~src_idx.index.isin(dest_idx.index)]],
        axis=0
    ).reset_index()

    try:
        combined["periodo_sort"] = pd.to_datetime(
            combined["periodo"] + "-01", errors="coerce"
        )
        combined = combined.sort_values(
            ["periodo_sort", "codigo_proyecto_salud", "codigo_proceso_recurso", "prot_codigo"]
        ).drop(columns=["periodo_sort"])
    except Exception:
        combined = combined.sort_values(
            ["periodo", "codigo_proyecto_salud", "codigo_proceso_recurso", "prot_codigo"]
        )

    combined["horas_recibidas_salud"] = combined["horas_recibidas_salud"].round(2)
    return combined

def main():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encuentra el Excel: {EXCEL_PATH}")

    df_src = read_excel_detalle(EXCEL_PATH, EXCEL_SHEET)

    gc = connect_gsheets(SERVICE_ACCOUNT_FILE)
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME_DEST)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME_DEST, rows=1000, cols=10)

    df_dest = read_ws_as_df(ws)
    df_final = upsert_destino(df_dest, df_src)
    write_df_to_ws(ws, df_final)

    print("[INFO] Periodos importados:", sorted(df_src["periodo"].unique())[:10])
    print(f"[OK] Hoja '{SHEET_NAME_DEST}' actualizada desde '{EXCEL_SHEET}'.")
    print(f"[INFO] Filas origen: {len(df_src)} | Filas destino final: {len(df_final)}")

if __name__ == "__main__":
    main()
