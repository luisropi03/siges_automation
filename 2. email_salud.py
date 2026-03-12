
import base64, re
from pathlib import Path
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ========= CONFIG =========
BASE_DIR         = Path(__file__).resolve().parent
EXCEL_PATH       = BASE_DIR / "reporte_siges_salud.xlsx"
SHEET_GENERAL    = "general"
SHEET_DETALLE    = "detallexproyecto"

CREDENTIALS_FILE = "client_secret_612599563025-ha630o27m2rlgg48gp9739f1eceogqde.apps.googleusercontent.com.json"
TOKEN_FILE       = "token.json"
SCOPES           = ["https://www.googleapis.com/auth/gmail.compose"]

TO  = ["ajimenez@soin.co.cr"]
CC  = ["ivette@soin.co.cr"]
BCC = []

OUTPUT_HTML      = BASE_DIR / "reporte_siges_salud_email.html"
MAX_DETAIL_ROWS  = 500

# ========= UTILIDADES =========
MESES_ES = ["enero","febrero","marzo","abril","mayo","junio","julio",
            "agosto","septiembre","octubre","noviembre","diciembre"]

def periodo_mes_anterior_es():
    hoy = datetime.now()
    primero_mes = datetime(hoy.year, hoy.month, 1)
    ultimo_mes_anterior = primero_mes - timedelta(days=1)
    mes = MESES_ES[ultimo_mes_anterior.month - 1]
    return f"{mes.capitalize()} {ultimo_mes_anterior.year}"

SUBJECT = f"[SIGES] SALUD – horas recibidas de procesos clave ({periodo_mes_anterior_es()})"

def fmt_horas(v) -> str:
    try:
        x = float(v)
        if abs(x) < 0.005:
            x = 0.0
        s = f"{x:,.2f}"
        return s.replace(",", "_").replace(".", ",").replace("_", ".")
    except Exception:
        return str(v)

def neutralize_proc_autolinks(html: str) -> str:
    return re.sub(r'(?i)PROC\.', 'PROC.&#8203;', html)

# ========= LECTURA =========
def leer_dataframes():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encontró el Excel: {EXCEL_PATH}")

    df_general = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_GENERAL, dtype={"code": str})
    df_general.columns = [str(c).strip() for c in df_general.columns]

    df_general["hours_otros_brindan_a_salud"] = pd.to_numeric(
        df_general["hours_otros_brindan_a_salud"], errors="coerce"
    ).fillna(0.0)

    total_horas_general = df_general["hours_otros_brindan_a_salud"].sum()
    top3 = df_general.sort_values("hours_otros_brindan_a_salud", ascending=False).head(3)

    df_detalle = pd.read_excel(
        EXCEL_PATH,
        sheet_name=SHEET_DETALLE,
        dtype={"PRYcodigo": str, "PRYTcodigo": str, "ProcesoRecurso": str}
    )
    df_detalle.columns = [str(c).strip() for c in df_detalle.columns]

    df_detalle["PRYRAhoras"] = pd.to_numeric(
        df_detalle["PRYRAhoras"], errors="coerce"
    ).fillna(0.0)

    if pd.api.types.is_numeric_dtype(df_detalle["PRYRAfecha"]):
        def yyyymm_to_date(v):
            try:
                v = int(v); y, m = divmod(v, 100)
                return datetime(year=y, month=m, day=1).date()
            except Exception:
                return pd.NaT
        df_detalle["PRYRAfecha"] = df_detalle["PRYRAfecha"].apply(yyyymm_to_date)
    else:
        df_detalle["PRYRAfecha"] = pd.to_datetime(
            df_detalle["PRYRAfecha"], errors="coerce"
        ).dt.date

    sum_por_pry = (
        df_detalle.groupby("PRYcodigo", as_index=False)["PRYRAhoras"]
        .sum()
        .rename(columns={"PRYRAhoras": "Total_PRYRAhoras"})
        .sort_values("Total_PRYRAhoras", ascending=False)
    )

    return df_general, df_detalle, total_horas_general, top3, sum_por_pry

# ========= HTML (ORIGINAL, INTACTO) =========
TH_STYLE = "border:1px solid #e5e7eb; padding:8px 10px; background:#f9fafb; text-align:left;"
TD_STYLE = "border:1px solid #e5e7eb; padding:8px 10px; text-align:left;"
TD_NUM_STYLE = "border:1px solid #e5e7eb; padding:8px 10px; text-align:right;"

def table_open(min_width=560):
    return f"<table style='border-collapse:collapse; min-width:{min_width}px;'>"

def tr_head(*cells):
    return "<tr>" + "".join(f"<th style='{TH_STYLE}'>{c}</th>" for c in cells) + "</tr>"

def tr_row(values, num_idx=None):
    num_idx = set(num_idx or [])
    tds = []
    for i, v in enumerate(values):
        style = TD_NUM_STYLE if i in num_idx else TD_STYLE
        tds.append(f"<td style='{style}'>{v}</td>")
    return "<tr>" + "".join(tds) + "</tr>"

def html_base_css():
    return """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; color:#1f2937; font-size:14px; }
    p  { margin: 6px 0; }
    ul { margin: 6px 0 10px 16px; }
    .section-title { margin: 12px 0 8px 0; font-weight:600; }
    .muted { color:#6b7280; font-size:12px; }
  </style>
</head>
<body>
"""

def html_header(periodo: str, total_general: float, top3: pd.DataFrame):
    items_top = []
    for _, r in top3.iterrows():
        items_top.append(f"{r['process_name']} ({r['code']}): {fmt_horas(r['hours_otros_brindan_a_salud'])} h")
    top_str = ", ".join(items_top) if items_top else "—"
    return f"""
  <p>Buen día, Don Adrián y Sra. Ivette.</p>
  <p>Comparto el resumen de <b>SALUD</b> con base en SIGES para <b>{periodo}</b>:</p>
  <ul>
    <li>Horas totales <b>recibidas por SALUD</b> (procesos objetivo): <b>{fmt_horas(total_general)}</b> h.</li>
    <li>Top 3 aportantes: {top_str}.</li>
  </ul>
"""

def html_tabla_general(df_general: pd.DataFrame):
    df = df_general.sort_values("hours_otros_brindan_a_salud", ascending=False).copy()
    rows_html = []
    for _, r in df.iterrows():
        rows_html.append(tr_row(
            [f"{r['process_name']} ({r['code']})", fmt_horas(r["hours_otros_brindan_a_salud"])],
            num_idx={1}
        ))
    total = fmt_horas(df["hours_otros_brindan_a_salud"].sum())
    return (
        "<br>"
        "<div class='section-title'>Detalle de horas brindadas a SALUD por proceso:</div>"
        "<br>"
        + table_open(560)
        + "<thead>" + tr_head("Proceso", "Horas recibidas por SALUD") + "</thead>"
        + "<tbody>"
        + "".join(rows_html)
        + tr_row(["<b>Total</b>", f"<b>{total}</b>"], num_idx={1})
        + "</tbody></table>"
    )

def html_tabla_detalle(df_detalle: pd.DataFrame):
    if df_detalle is None or df_detalle.empty:
        return ""
    df_all = df_detalle.sort_values(["PRYcodigo","ProcesoRecurso","PRYRAfecha"]).copy()
    df = df_all if len(df_all) <= MAX_DETAIL_ROWS else df_all.head(MAX_DETAIL_ROWS)
    body = []
    for _, r in df.iterrows():
        body.append(tr_row(
            [
                r['PRYcodigo'],
                r['ProcesoRecurso'],
                r['PRYTcodigo'],
                "" if pd.isna(r['PRYRAfecha']) else r['PRYRAfecha'],
                fmt_horas(r['PRYRAhoras'])
            ],
            num_idx={4}
        ))
    total = fmt_horas(df['PRYRAhoras'].sum())
    return (
        "<br>"
        "<div class='section-title'>Detalle por proyecto de SALUD:</div>"
        "<br>"
        + table_open(560)
        + "<thead>" + tr_head("Codigo de Proyecto SALUD", "Codigo de Proceso-Recurso", "PRYTcodigo", "Periodo", "Horas recibidas por SALUD") + "</thead>"
        + "<tbody>"
        + "".join(body)
        + tr_row(["<b>Total</b>", "", "", "", f"<b>{total}</b>"], num_idx={4})
        + "</tbody></table>"
    )

def html_tabla_sumatoria_por_pry(sum_por_pry: pd.DataFrame):
    if sum_por_pry is None or sum_por_pry.empty:
        return ""
    body = []
    for _, r in sum_por_pry.iterrows():
        body.append(tr_row([r["PRYcodigo"], fmt_horas(r["Total_PRYRAhoras"])], num_idx={1}))
    total = fmt_horas(sum_por_pry["Total_PRYRAhoras"].sum())
    return (
        "<br>"
        "<div class='section-title'>Sumatoria de horas recibidas por Codigo de Proyecto SALUD</div>"
        "<br>"
        + table_open(560)
        + "<thead>" + tr_head("Codigo de Proyecto SALUD", "Horas recibidas por SALUD") + "</thead>"
        + "<tbody>"
        + "".join(body)
        + tr_row(["<b>Total</b>", f"<b>{total}</b>"], num_idx={1})
        + "</tbody></table>"
    )

def html_footer():
    return """
  <p class="muted">
    <p> Adjunto documento con el detalle de horas.</p>
    Notas: Las cifras provienen del archivo 'Indicadores por Proceso V1' tras refrescar el archivo con el parámetro <b>PROC.SALUD</b>.
  </p>
  <p>— Reporte SIGES</p>
</body>
</html>
"""

def construir_html(df_general, df_detalle, total_horas_general, top3, sum_por_pry):
    periodo = periodo_mes_anterior_es()
    parts = []
    parts.append(html_base_css())
    parts.append(html_header(periodo, total_horas_general, top3))
    parts.append(html_tabla_general(df_general))
    parts.append("<br><br>")
    parts.append(html_tabla_detalle(df_detalle))
    parts.append("<br><br>")
    parts.append(html_tabla_sumatoria_por_pry(sum_por_pry))
    parts.append(html_footer())
    html = "".join(parts)
    return neutralize_proc_autolinks(html)

# ========= ADJUNTO =========
def get_detalle_horas_file() -> Path:
    hoy = datetime.now()
    primero_mes = datetime(hoy.year, hoy.month, 1)
    ultimo_mes_anterior = primero_mes - timedelta(days=1)
    fname = f"Detalle de horas SALUD {ultimo_mes_anterior.month:02d}{ultimo_mes_anterior.year}.xlsx"
    path = BASE_DIR / fname
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo adjunto: {path}")
    return path

def attach_file(msg: MIMEMultipart, file_path: Path):
    part = MIMEBase("application", "octet-stream")
    part.set_payload(file_path.read_bytes())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
    msg.attach(part)

# ========= GMAIL =========
def get_service():
    creds = None
    token_path = Path(TOKEN_FILE)
    cred_path = BASE_DIR / CREDENTIALS_FILE
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    return build("gmail", "v1", credentials=creds)

def build_mime_message(html_body: str, attachment: Path | None = None) -> MIMEMultipart:
    msg = MIMEMultipart()
    msg["Subject"] = SUBJECT
    msg["To"] = ", ".join(TO)
    if CC:  msg["Cc"]  = ", ".join(CC)
    if BCC: msg["Bcc"] = ", ".join(BCC)
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    if attachment:
        attach_file(msg, attachment)
    return msg

def create_gmail_draft(service, mime_msg: MIMEMultipart):
    raw = base64.urlsafe_b64encode(mime_msg.as_bytes()).decode("utf-8")
    body = {"message": {"raw": raw}}
    return service.users().drafts().create(userId="me", body=body).execute()

# ========= MAIN =========
def main():
    df_general, df_detalle, total_g, top3, sum_por_pry = leer_dataframes()
    html_body = construir_html(df_general, df_detalle, total_g, top3, sum_por_pry)

    OUTPUT_HTML.write_text(html_body, encoding="utf-8")

    detalle_file = get_detalle_horas_file()

    service = get_service()
    draft = create_gmail_draft(
        service,
        build_mime_message(html_body, attachment=detalle_file)
    )

    print("[OK] Borrador creado en Gmail.")
    print("[INFO] Draft ID:", draft.get("id"))
    print(f"[INFO] Adjunto: {detalle_file.name}")
    print(f"[INFO] HTML guardado en: {OUTPUT_HTML}")

if __name__ == "__main__":
    main()
