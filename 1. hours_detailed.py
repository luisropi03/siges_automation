# hours_detailed.py
import sys, time
import re
from pathlib import Path
from datetime import date, timedelta
import pythoncom
import pandas as pd
from win32com.client import Dispatch, DispatchEx, constants

# ========= CONFIG =========
FILENAME = "Copia Reportes Siges v2.2 3.xlsm"

SHEET_CONTROL = "Control"
SHEET_HORAS   = "Horas"

VALUES_ROW_17 = {
    "F17": "095-202-RDIG",
    "G17": "151-400-RGT2.0",
    "H17": "579-203-REDIMED",
    "I17": "095-192-RGT",
    "J17": "151-350-Dicomizer",
    "K17": "095-400-INS-REG",
}

CELL_START_DATE = "F3"
CELL_END_DATE   = "F4"

HORAS_COLUMNS = [
    "PRYTcodigo",
    "PRYcodigo",
    "PRYdescripcion",
    "PRYAdescripcion",
    "PRYEcodigo",
    "PRYEdescripcion",
    "Usulogin",
    "Usunombre",
    "PRYRAfecha",
    "PRYRAcomentario",
    "PRYRAhoras",
]

REFRESH_TIMEOUT_S = 900
RETRY_SLEEP_S = 0.5
RETRY_MAX = 120
RPC_E_CALL_REJECTED = -2147418111
LOGIN_SESSION_ERROR = -2147023584
ILLEGAL_XLSX_TEXT_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")

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
        excel.EnableEvents = True
        excel.AutomationSecurity = constants.msoAutomationSecurityLow
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
            raise TimeoutError(f"Refresh excedio {REFRESH_TIMEOUT_S} segundos")

        time.sleep(1)

# ========= Helpers =========
def used_range_to_dataframe(ws) -> pd.DataFrame:
    values = ws.UsedRange.Value
    if not values:
        return pd.DataFrame()

    rows = [list(r) if isinstance(r, (list, tuple)) else [r] for r in values]
    rows = [r for r in rows if any(x is not None and str(x).strip() != "" for x in r)]
    if not rows:
        return pd.DataFrame()

    headers = [str(h).strip() for h in rows[0]]
    data = rows[1:]

    # 👇 CLAVE: TODO entra como TEXTO
    df = pd.DataFrame(data, columns=headers, dtype=object)
    df = df.apply(lambda col: col.map(lambda x: "" if x is None else str(x).strip()))

    return df[[c for c in df.columns if c]]

def sanitize_for_xlsx(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina caracteres de control no permitidos por XLSX en columnas de texto."""
    out = df.copy()
    text_cols = out.select_dtypes(include=["object"]).columns
    for c in text_cols:
        out[c] = out[c].map(
            lambda v: ILLEGAL_XLSX_TEXT_RE.sub("", v) if isinstance(v, str) else v
        )
    return out

# ========= Fechas =========
def previous_month_first_and_last():
    today = date.today()
    first_current = today.replace(day=1)
    last_prev = first_current - timedelta(days=1)
    first_prev = last_prev.replace(day=1)
    return first_prev, last_prev

def fmt_mmddyyyy(d: date) -> str:
    return d.strftime("%m/%d/%Y")

# ========= Lectura Horas =========
def read_horas_df(wb) -> pd.DataFrame:
    ws = wb.Worksheets(SHEET_HORAS)
    df = used_range_to_dataframe(ws)

    cols_lower = {c.lower(): c for c in df.columns}
    df = df[[cols_lower[c.lower()] for c in HORAS_COLUMNS]].copy()

    # 👇 Todo queda como texto EXCEPTO horas
    for c in df.columns:
        if c != "PRYRAhoras":
            df[c] = df[c].astype(str).str.strip()

    # 👇 ÚNICA conversión numérica (segura)
    df["PRYRAhoras"] = (
        pd.to_numeric(df["PRYRAhoras"], errors="coerce")
        .fillna(0.0)
    )

    df = df[df["PRYRAhoras"] > 0]
    return df.reset_index(drop=True)

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
        wb = excel.Workbooks.Open(str(file_path), ReadOnly=False)
        ws = wb.Worksheets(SHEET_CONTROL)

        for cell, value in VALUES_ROW_17.items():
            ws.Range(cell).Value = value

        first_prev, last_prev = previous_month_first_and_last()
        ws.Range(CELL_START_DATE).Value = fmt_mmddyyyy(first_prev)
        ws.Range(CELL_END_DATE).Value   = fmt_mmddyyyy(last_prev)

        print("[INFO] Ejecutando consulta...")
        excel.Run("'Copia Reportes Siges v2.2 3.xlsm'!Módulo1.Query")

        refresh_all(excel)

        print("[INFO] Extrayendo horas...")
        df = read_horas_df(wb)
        df = sanitize_for_xlsx(df)

        out_file = base / f"Detalle de horas SALUD {first_prev.month:02d}{first_prev.year}.xlsx"
        df.to_excel(out_file, index=False)

        print(f"[OK] Archivo generado: {out_file.name}")

        wb.Close(SaveChanges=True)
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
