# key_processes_extration.py
import sys, time
from pathlib import Path
from datetime import date, timedelta
import pythoncom
import pandas as pd
from win32com.client import Dispatch, DispatchEx, constants

# ========= CONFIG =========
FILENAME         = "indicadores.xlsx"

# Parámetro que dispara el refresh
SHEET_PARAM      = "Resumen"
CELL_PARAM       = "D7"
PROC_VALUE       = "PROC.SALUD"

# Hoja/zonas de interés (Global)
SHEET_GLOBAL     = "Global"
COL_HEADER_CELL  = "R1"
ROW_LABELS_RG    = "B2:B32"
COL_VALUES_RG    = "R2:R32"

# Hoja BD (detalle)
SHEET_BD         = "BD"

# Filtro correcto (según lo que pedís ahora): por ProcesoTarea
BD_FILTER_PROCESO_TAREA = "PROC.SALUD"

BD_OUT_COLUMNS = ["PRYTcodigo", "PRYcodigo", "PRYRAfecha", "PRYRAhoras", "ProcesoRecurso"]

# Archivo destino
REPORT_FILENAME  = "reporte_siges_salud.xlsx"
REPORT_SHEET_OUT_GENERAL = "general"
REPORT_SHEET_OUT_DETALLE = "detallexproyecto"

TARGET_CODES = [
    "PROC.CUST", "PROC.VENTA", "PROC.POOL", "PROC.SWAT",
    "PROC.QA", "PROC.RM", "PROC.PMO", "PROC.BI",
]
TARGET_CODES_SET = {c.upper() for c in TARGET_CODES}

PROCESS_NAMES = {
    "PROC.CUST":  "Proceso Customer Success",
    "PROC.VENTA": "Proceso Comercial",
    "PROC.POOL":  "Proceso Pool de Integraciones",
    "PROC.SWAT":  "Equipo SWAT",
    "PROC.QA":    "Proceso Aseguramiento de la Calidad",
    "PROC.RM":    "Proceso Liberacion de Software",
    "PROC.PMO":   "Proceso Gestion de Proyectos",
    "PROC.BI":    "Proceso Inteligencia de Negocios",
}

REFRESH_TIMEOUT_S = 900
RETRY_SLEEP_S     = 0.5
RETRY_MAX         = 120
RPC_E_CALL_REJECTED = -2147418111
LOGIN_SESSION_ERROR = -2147023584

# ========= Helpers COM =========
def com_call(fn, *args, **kwargs):
    tries = 0
    while True:
        try:
            return fn(*args, **kwargs)
        except pythoncom.com_error as e:
            if getattr(e, "hresult", None) == RPC_E_CALL_REJECTED and tries < RETRY_MAX:
                tries += 1
                time.sleep(RETRY_SLEEP_S)
                continue
            raise

def get_prop(obj, prop_name):
    tries = 0
    while True:
        try:
            return getattr(obj, prop_name)
        except pythoncom.com_error as e:
            if getattr(e, "hresult", None) == RPC_E_CALL_REJECTED and tries < RETRY_MAX:
                tries += 1
                time.sleep(RETRY_SLEEP_S)
                continue
            raise

def excel_setup():
    pythoncom.CoInitialize()
    last_exc = None

    for creator in (Dispatch, DispatchEx):
        try:
            excel = creator("Excel.Application")
            break
        except pythoncom.com_error as e:
            last_exc = e
            excel = None
    else:
        if getattr(last_exc, "hresult", None) == LOGIN_SESSION_ERROR:
            raise RuntimeError(
                "No se puede iniciar Excel por COM porque no hay una sesion de Windows activa "
                "para este usuario. Ejecuta el script con sesion interactiva iniciada."
            )
        raise last_exc

    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AutomationSecurity = constants.msoAutomationSecurityForceDisable
        excel.Calculation = constants.xlCalculationManual
    except Exception:
        pass
    return excel

def refresh_all(excel):
    try:
        com_call(excel.CalculateUntilAsyncQueriesDone)
    except Exception:
        pass

    start = time.time()
    while True:
        try:
            refreshing = bool(getattr(excel, "RefreshingData", False))
        except pythoncom.com_error as e:
            if getattr(e, "hresult", None) == RPC_E_CALL_REJECTED:
                time.sleep(RETRY_SLEEP_S)
                continue
            raise

        if not refreshing:
            break

        if time.time() - start > REFRESH_TIMEOUT_S:
            raise TimeoutError(f"RefreshAll excedio {REFRESH_TIMEOUT_S} segundos")

        time.sleep(1)

# ========= Helpers DF =========
def to_col_list(v):
    if v is None:
        return []
    if isinstance(v, (list, tuple)):
        if v and isinstance(v[0], (list, tuple)):
            return [row[0] if row else None for row in v]
        return list(v)
    return [v]

def used_range_to_dataframe(ws) -> pd.DataFrame:
    used = get_prop(ws, "UsedRange")
    values = get_prop(used, "Value")
    if values is None:
        return pd.DataFrame()

    rows = [list(r) if isinstance(r, (list, tuple)) else [r] for r in values]
    rows = [r for r in rows if any(x is not None and str(x).strip() != "" for x in r)]
    if not rows:
        return pd.DataFrame()

    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = rows[1:] if len(rows) > 1 else []
    df = pd.DataFrame(data, columns=headers)
    df = df[[c for c in df.columns if str(c).strip() != ""]]
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalize_str_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip()

# ========= Fechas =========
def previous_month_yyyymm(today: date | None = None) -> int:
    """Devuelve YYYYMM del mes anterior (ej: 202512)."""
    today = today or date.today()
    first_current = today.replace(day=1)
    last_prev = first_current - timedelta(days=1)
    return int(f"{last_prev.year}{last_prev.month:02d}")

def to_yyyymm_any(v) -> int | None:
    """Normaliza PRYRAfecha a int YYYYMM."""
    if v is None:
        return None

    if isinstance(v, int):
        s = str(v)
    elif isinstance(v, float):
        if pd.isna(v):
            return None
        if abs(v) > 100000:
            s = str(int(v))
        else:
            s = str(v).strip()
    else:
        s = str(v).strip()

    if not s:
        return None

    if s.isdigit() and len(s) == 6:
        return int(s)
    if s.isdigit() and len(s) == 8:
        return int(s[:6])

    if s.isdigit() and len(s) <= 5:
        try:
            serial = int(s)
            dt = pd.to_datetime(serial, unit="D", origin="1899-12-30", errors="coerce")
            if pd.isna(dt):
                return None
            return int(f"{dt.year}{dt.month:02d}")
        except Exception:
            return None

    dt = pd.to_datetime(s, errors="coerce")
    if pd.isna(dt):
        return None
    return int(f"{dt.year}{dt.month:02d}")

# ========= Negocio =========
def build_global_df(wb):
    ws = com_call(wb.Worksheets, SHEET_GLOBAL)
    col_header = str(com_call(ws.Range, COL_HEADER_CELL).Value).strip().upper()

    if col_header != PROC_VALUE:
        raise RuntimeError(
            f"{SHEET_GLOBAL}!{COL_HEADER_CELL} = '{col_header}', esperado '{PROC_VALUE}'"
        )

    labels = [("" if x is None else str(x).strip())
              for x in to_col_list(com_call(ws.Range, ROW_LABELS_RG).Value)]
    vals = to_col_list(com_call(ws.Range, COL_VALUES_RG).Value)

    df = pd.DataFrame({
        "code": labels[:30],
        "hours_otros_brindan_a_salud": pd.to_numeric(vals[:30], errors="coerce")
    })

    df["code"] = df["code"].str.upper().str.strip()
    df = df[df["code"].isin(TARGET_CODES_SET)].copy()
    df["process_name"] = df["code"].map(PROCESS_NAMES)

    order = ["PROC.CUST","PROC.VENTA","PROC.POOL","PROC.SWAT","PROC.QA","PROC.RM","PROC.PMO","PROC.BI"]
    df["__ord"] = df["code"].apply(lambda x: order.index(x) if x in order else 999)
    df = df.sort_values("__ord").drop(columns="__ord").reset_index(drop=True)

    return df[["code", "process_name", "hours_otros_brindan_a_salud"]]

def build_bd_detalle_df(wb) -> pd.DataFrame:
    target_period = previous_month_yyyymm()

    ws = com_call(wb.Worksheets, SHEET_BD)
    df_bd = used_range_to_dataframe(ws)
    if df_bd.empty:
        return df_bd

    cols_lower = {c.lower(): c for c in df_bd.columns}

    # OJO: ahora requerimos ProcesoTarea para filtrar correctamente
    required = ["prytcodigo", "prycodigo", "pryrafecha", "pryrahoras", "procesorecurso", "procesotarea"]
    missing = [c for c in required if c not in cols_lower]
    if missing:
        raise ValueError(f"Faltan columnas en '{SHEET_BD}': {missing}")

    df = df_bd[[cols_lower[c] for c in required]].copy()
    df.columns = ["PRYTcodigo", "PRYcodigo", "PRYRAfecha", "PRYRAhoras", "ProcesoRecurso", "ProcesoTarea"]

    # Normalizar strings
    df["PRYTcodigo"] = normalize_str_series(df["PRYTcodigo"]).str.upper()
    df["PRYcodigo"] = normalize_str_series(df["PRYcodigo"])
    df["ProcesoRecurso"] = normalize_str_series(df["ProcesoRecurso"]).str.upper()
    df["ProcesoTarea"] = normalize_str_series(df["ProcesoTarea"]).str.upper()

    # Periodo YYYYMM
    df["PRYRAfecha"] = df["PRYRAfecha"].apply(to_yyyymm_any)

    # Horas numéricas
    df["PRYRAhoras"] = pd.to_numeric(df["PRYRAhoras"], errors="coerce").fillna(0.0)

    # Filtros mínimos
    # ANTES: df = df[df["PRYTcodigo"] == "PROC.SALUD"]
    # AHORA: filtrar por ProcesoTarea
    df = df[df["ProcesoTarea"] == BD_FILTER_PROCESO_TAREA.upper()].copy()
    df = df[df["PRYRAfecha"] == target_period].copy()

    # Solo procesos objetivo (por si BD trae otros)
    df = df[df["ProcesoRecurso"].isin(TARGET_CODES_SET)].copy()

    # Agrupar por PRYcodigo + ProcesoRecurso (y mantener PRYTcodigo + PRYRAfecha)
    df = (
        df.groupby(["PRYTcodigo", "PRYcodigo", "PRYRAfecha", "ProcesoRecurso"], as_index=False)["PRYRAhoras"]
        .sum()
    )

    df = df[BD_OUT_COLUMNS].sort_values(["PRYcodigo", "ProcesoRecurso"]).reset_index(drop=True)
    return df

def write_to_report(base_dir: Path, df_general: pd.DataFrame, df_detalle: pd.DataFrame):
    report_path = base_dir / REPORT_FILENAME
    if not report_path.exists():
        raise FileNotFoundError(f"No se encontro el archivo destino: {report_path}")

    with pd.ExcelWriter(report_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_general.to_excel(writer, sheet_name=REPORT_SHEET_OUT_GENERAL, index=False)
        df_detalle.to_excel(writer, sheet_name=REPORT_SHEET_OUT_DETALLE, index=False)

    print(f"[OK] Hojas '{REPORT_SHEET_OUT_GENERAL}' y '{REPORT_SHEET_OUT_DETALLE}' actualizadas en {REPORT_FILENAME}")

# ========= MAIN =========
def main():
    base = Path(__file__).resolve().parent
    file_path = base / FILENAME

    if not file_path.exists():
        print(f"[ERROR] No se encontro el archivo: {file_path}")
        return 1

    excel = excel_setup()
    wb = None

    try:
        wb = com_call(excel.Workbooks.Open, str(file_path), ReadOnly=False)

        ws_param = com_call(wb.Worksheets, SHEET_PARAM)
        com_call(ws_param.Range, CELL_PARAM).__setattr__("Value", PROC_VALUE)
        print(f"[OK] Escrito {SHEET_PARAM}!{CELL_PARAM} = '{PROC_VALUE}'")

        print("[INFO] Refrescando conexiones...")
        com_call(wb.RefreshAll)
        refresh_all(excel)

        df_general = build_global_df(wb)
        print("[OK] DataFrame 'general' construido")

        target_period = previous_month_yyyymm()
        df_detalle = build_bd_detalle_df(wb)
        print(f"[OK] DataFrame 'detallexproyecto' construido (periodo {target_period}, filas: {len(df_detalle)})")

        write_to_report(base, df_general, df_detalle)

        wb.Close(SaveChanges=False)
        excel.Quit()
        print("[OK] Excel cerrado correctamente")
        return 0

    except Exception as e:
        try:
            wb and wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        print(f"[ERROR] {e}")
        return 1
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

if __name__ == "__main__":
    sys.exit(main())
