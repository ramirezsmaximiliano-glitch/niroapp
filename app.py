import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
import json
import hashlib
import base64
from PIL import Image
import requests
from pathlib import Path

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="NIRO — Control de Efectivo",
    page_icon="🧡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

ADMIN_PASSWORD = "Niro26"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Nunito:wght@400;600;700;800&display=swap');

/* Reset & base */
html, body, [class*="css"] {
    font-family: 'Nunito', sans-serif;
    background-color: #0f0f0f;
    color: #f0f0f0;
}

/* Header NIRO */
.niro-header {
    background: linear-gradient(135deg, #1a1a1a 0%, #0f0f0f 100%);
    border-bottom: 3px solid #E8450A;
    padding: 1.5rem 2rem;
    margin: -1rem -1rem 2rem -1rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.niro-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.4rem;
    color: #ffffff;
    letter-spacing: 2px;
    line-height: 1.1;
    margin: 0;
}
.niro-subtitle {
    font-size: 0.75rem;
    color: #E8450A;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin: 0;
}

/* Cards */
.card {
    background: #1a1a1a;
    border: 1px solid #2a2a2a;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
    border-left: 4px solid #E8450A;
}
.card-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.1rem;
    letter-spacing: 1.5px;
    color: #E8450A;
    margin-bottom: 0.5rem;
}

/* Botones */
div.stButton > button {
    background: #E8450A;
    color: white;
    border: none;
    border-radius: 8px;
    font-family: 'Nunito', sans-serif;
    font-weight: 800;
    font-size: 1rem;
    padding: 0.6rem 2rem;
    width: 100%;
    cursor: pointer;
    transition: all 0.2s ease;
    letter-spacing: 0.5px;
}
div.stButton > button:hover {
    background: #c73a08;
    transform: translateY(-1px);
    box-shadow: 0 4px 15px rgba(232, 69, 10, 0.4);
}

/* Inputs */
div[data-testid="stSelectbox"] label,
div[data-testid="stTextInput"] label,
div[data-testid="stNumberInput"] label,
div[data-testid="stTextArea"] label,
div[data-testid="stFileUploader"] label {
    font-weight: 700;
    color: #cccccc;
    font-size: 0.85rem;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}

/* Success / error */
.success-box {
    background: #0d2b1a;
    border: 1px solid #1e7a45;
    border-left: 4px solid #2ecc71;
    border-radius: 8px;
    padding: 1rem 1.5rem;
    color: #2ecc71;
    font-weight: 700;
    font-size: 1rem;
}
.error-box {
    background: #2b0d0d;
    border: 1px solid #7a1e1e;
    border-left: 4px solid #e74c3c;
    border-radius: 8px;
    padding: 1rem 1.5rem;
    color: #e74c3c;
    font-weight: 700;
}

/* Folio badge */
.folio-badge {
    display: inline-block;
    background: #E8450A;
    color: white;
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.3rem;
    padding: 0.3rem 1rem;
    border-radius: 6px;
    letter-spacing: 2px;
    margin-bottom: 0.5rem;
}

/* Admin panel */
.admin-section {
    background: #1a1a1a;
    border: 1px solid #333;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
}

/* Step indicator */
.step-badge {
    display: inline-block;
    background: #E8450A;
    color: white;
    font-family: 'Bebas Neue', sans-serif;
    font-size: 0.9rem;
    padding: 0.15rem 0.6rem;
    border-radius: 4px;
    margin-right: 0.5rem;
    letter-spacing: 1px;
}

/* Divider */
.divider {
    border: none;
    border-top: 1px solid #2a2a2a;
    margin: 1.5rem 0;
}

/* Semana tag */
.semana-tag {
    background: #2a2a2a;
    color: #E8450A;
    font-size: 0.75rem;
    font-weight: 700;
    padding: 0.2rem 0.6rem;
    border-radius: 4px;
    letter-spacing: 1px;
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# HELPERS: GOOGLE
# ─────────────────────────────────────────────
def get_google_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client, creds

def get_drive_service(creds):
    return build("drive", "v3", credentials=creds)

def get_or_create_sheet(client, spreadsheet_id):
    return client.open_by_key(spreadsheet_id)

def get_week_sheet_name():
    today = datetime.now()
    # Lunes de esta semana
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    week_num = monday.isocalendar()[1]
    year = monday.year
    return f"Semana {week_num} ({monday.strftime('%d/%m')} - {sunday.strftime('%d/%m/%Y')})"

def ensure_week_sheet(spreadsheet):
    sheet_name = get_week_sheet_name()
    sheet_titles = [ws.title for ws in spreadsheet.worksheets()]
    
    if sheet_name not in sheet_titles:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=500, cols=10)
        headers = ["Folio", "Fecha y Hora", "Rider", "Agencia", "Monto (MXN)", "Recibido por", "Notas", "Comprobante URL"]
        ws.append_row(headers)
        
        # Formato encabezado
        ws.format("A1:H1", {
            "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.067, "green": 0.067, "blue": 0.067},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
        })
        
        # Ancho de columnas
        spreadsheet.batch_update({"requests": [
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 160}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
                "properties": {"pixelSize": 145}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
                "properties": {"pixelSize": 160}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4},
                "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5},
                "properties": {"pixelSize": 120}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 6},
                "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 6, "endIndex": 7},
                "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 8},
                "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
            # Altura del header
            {"updateDimensionProperties": {
                "range": {"sheetId": ws.id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 35}, "fields": "pixelSize"}},
            # Freeze header row
            {"updateSheetProperties": {
                "properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount"}},
        ]})
        
        return ws, True
    else:
        ws = spreadsheet.worksheet(sheet_name)
        return ws, False


def format_data_row(ws, spreadsheet, row_num):
    """Aplica formato a una fila de datos recién agregada"""
    # Alternar color de fila
    bg = {"red": 0.95, "green": 0.95, "blue": 0.95} if row_num % 2 == 0 else {"red": 1, "green": 1, "blue": 1}
    
    ws.format(f"A{row_num}:H{row_num}", {
        "backgroundColor": bg,
        "verticalAlignment": "MIDDLE",
        "textFormat": {"fontSize": 10},
    })
    # Formato dinero en columna E
    ws.format(f"E{row_num}", {
        "numberFormat": {"type": "CURRENCY", "pattern": '$#,##0.00'},
        "horizontalAlignment": "RIGHT",
    })
    # Centrar columnas A, B, D, F
    ws.format(f"A{row_num}:D{row_num}", {"horizontalAlignment": "CENTER"})
    ws.format(f"F{row_num}", {"horizontalAlignment": "CENTER"})

def get_next_folio(ws):
    records = ws.get_all_values()
    # Contar filas con datos (excl header)
    count = len([r for r in records[1:] if r[0]])
    return f"NIRO-{datetime.now().strftime('%y%m%d')}-{str(count + 1).zfill(3)}"

def append_registro(ws, spreadsheet, folio, fecha_hora, rider, agencia, monto, recibido_por, notas, comprobante_url):
    row = [folio, fecha_hora, rider, agencia, monto, recibido_por, notas, comprobante_url]
    ws.append_row(row)
    # Aplicar formato a la nueva fila
    all_rows = ws.get_all_values()
    row_num = len(all_rows)
    format_data_row(ws, spreadsheet, row_num)

def upload_comprobante_imgbb(file_bytes, filename):
    import base64
    api_key = "fa54f82ca6f8797309e7cf2ea06d21e6"
    b64 = base64.b64encode(file_bytes).decode("utf-8")
    response = requests.post(
        "https://api.imgbb.com/1/upload",
        data={"key": api_key, "image": b64, "name": filename}
    )
    data = response.json()
    if data.get("success"):
        return data["data"]["url_viewer"]
    return ""

def update_weekly_report(ws, spreadsheet_id, client, creds):
    """Genera resumen semanal al final de la hoja"""
    records = ws.get_all_values()
    if len(records) < 2:
        return
    
    df = pd.DataFrame(records[1:], columns=records[0])
    df = df[df["Folio"] != ""]
    df["Monto (MXN)"] = pd.to_numeric(df["Monto (MXN)"], errors="coerce").fillna(0)
    
    total = df["Monto (MXN)"].sum()
    por_rider = df.groupby("Rider")["Monto (MXN)"].sum().reset_index()
    
    # Fila donde empieza el resumen (dejar 2 filas de espacio)
    start_row = len(records) + 3
    
    # Limpiar resumen anterior si existe
    # Escribir resumen
    ws.update([["─── RESUMEN SEMANAL ───"]], f"A{start_row}")
    ws.update([["TOTAL GENERAL", "", f"${total:,.2f} MXN"]], f"A{start_row+1}")
    ws.update([["Rider", "Total (MXN)", ""]], f"A{start_row+2}")
    
    rider_rows = []
    for _, row in por_rider.iterrows():
        rider_rows.append([row["Rider"], f"${row['Monto (MXN)']:,.2f}", ""])
    ws.update(rider_rows, f"A{start_row+3}")

# ─────────────────────────────────────────────
# HELPERS: CONFIG (en Sheets — hoja "Config")
# ─────────────────────────────────────────────
def get_config_sheet(spreadsheet):
    titles = [ws.title for ws in spreadsheet.worksheets()]
    if "Config" not in titles:
        ws = spreadsheet.add_worksheet(title="Config", rows=100, cols=5)
        # Defaults
        ws.update([["RIDERS"]], "A1")
        default_riders = [
            ["Christian Alejandro"], ["Alan Daniel"], ["Carlos Omar"],
            ["Jair Asael"], ["David Martinez"], ["Felix de Jesus"], ["Ricardo Guadalupe"]
        ]
        ws.update(default_riders, "A2")
        
        ws.update([["AGENCIAS Y ENCARGADOS"]], "D1")
        ws.update([
            ["CF Leones", "Por definir"],
            ["Sierra Lincoln", "Por definir"],
        ], "D2")
        return ws
    return spreadsheet.worksheet("Config")

def load_config(spreadsheet):
    ws = get_config_sheet(spreadsheet)
    data = ws.get_all_values()
    
    riders = []
    agencias = {}
    
    for row in data[1:]:
        if row[0] and row[0] != "RIDERS":
            riders.append(row[0])
        if len(row) >= 4 and row[3] and row[3] not in ("AGENCIAS Y ENCARGADOS", ""):
            agencia = row[3]
            encargado = row[4] if len(row) > 4 else ""
            if agencia not in agencias:
                agencias[agencia] = []
            if encargado:
                agencias[agencia].append(encargado)
    
    return riders, agencias

def save_config(spreadsheet, riders, agencias):
    ws = get_config_sheet(spreadsheet)
    ws.clear()
    
    ws.update([["RIDERS"]], "A1")
    rider_rows = [[r] for r in riders]
    if rider_rows:
        ws.update(rider_rows, "A2")
    
    ws.update([["AGENCIAS Y ENCARGADOS"]], "D1")
    ag_rows = []
    for agencia, encargados in agencias.items():
        if encargados:
            for enc in encargados:
                ag_rows.append([agencia, enc])
        else:
            ag_rows.append([agencia, ""])
    if ag_rows:
        ws.update(ag_rows, "D2")


# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "admin_logged" not in st.session_state:
    st.session_state.admin_logged = False
if "registro_exitoso" not in st.session_state:
    st.session_state.registro_exitoso = None
if "vista" not in st.session_state:
    st.session_state.vista = "agencia"  # "agencia" | "admin"


# ─────────────────────────────────────────────
# LOGO + HEADER
# ─────────────────────────────────────────────
logo_b64 = ""
try:
    logo_path = Path("logo.png")
    if logo_path.exists():
        with open(logo_path, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
except:
    pass

logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:44px;" />' if logo_b64 else \
    '<span style="font-family:Bebas Neue,sans-serif;font-size:2rem;color:#E8450A;letter-spacing:3px;">NIRO</span>'

week_name = get_week_sheet_name()

st.markdown(f"""
<div class="niro-header">
    {logo_html}
    <div>
        <p class="niro-title">Control de Efectivo</p>
        <p class="niro-subtitle">DiDi Food · Agencias</p>
    </div>
    <div style="margin-left:auto; text-align:right;">
        <span class="semana-tag">{week_name}</span>
    </div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# NAVIGATION
# ─────────────────────────────────────────────
col1, col2 = st.columns([3, 1])
with col2:
    if st.session_state.vista == "agencia":
        if st.button("🔐 Admin"):
            st.session_state.vista = "admin"
            st.rerun()
    else:
        if st.button("← Volver"):
            st.session_state.vista = "agencia"
            st.session_state.admin_logged = False
            st.rerun()


# ─────────────────────────────────────────────
# VISTA: AGENCIA — FORMULARIO DE REGISTRO
# ─────────────────────────────────────────────
if st.session_state.vista == "agencia":

    # Mostrar éxito si acaba de registrar
    if st.session_state.registro_exitoso:
        folio = st.session_state.registro_exitoso
        st.markdown(f"""
        <div class="success-box">
            ✅ ¡Registro exitoso!<br>
            <span class="folio-badge">{folio}</span><br>
            <span style="font-size:0.85rem; color:#aaa;">Guarda este folio como comprobante.</span>
        </div>
        """, unsafe_allow_html=True)
        if st.button("➕ Nuevo Registro"):
            st.session_state.registro_exitoso = None
            st.rerun()
        st.stop()

    st.markdown('<div class="card"><p class="card-title">📦 Registro de Entrega de Efectivo</p>', unsafe_allow_html=True)
    st.markdown("Completa los datos de la entrega y sube el comprobante.", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Conectar Google
    try:
        client, creds = get_google_client()
        spreadsheet_id = st.secrets["spreadsheet_id"]
        folder_id = st.secrets["drive_folder_id"]
        spreadsheet = get_or_create_sheet(client, spreadsheet_id)
        riders, agencias = load_config(spreadsheet)
    except Exception as e:
        st.error(f"Error de conexión con Google: {e}")
        st.stop()

    agencia_lista = list(agencias.keys())

    with st.form("registro_form", clear_on_submit=False):
        st.markdown('<span class="step-badge">1</span> **¿Quién entrega?**', unsafe_allow_html=True)
        rider = st.selectbox("Rider", options=riders, key="rider_sel")
        
        st.markdown('<hr class="divider">', unsafe_allow_html=True)
        st.markdown('<span class="step-badge">2</span> **¿En qué agencia?**', unsafe_allow_html=True)
        
        agencia = st.selectbox("Agencia", options=agencia_lista, key="agencia_sel")
        
        # Encargados dinámicos según agencia seleccionada
        encargados_agencia = agencias.get(agencia, [])
        if encargados_agencia:
            recibido_por = st.selectbox("Recibido por", options=encargados_agencia)
        else:
            recibido_por = st.text_input("Recibido por (escribe el nombre)")
        
        st.markdown('<hr class="divider">', unsafe_allow_html=True)
        st.markdown('<span class="step-badge">3</span> **Monto entregado**', unsafe_allow_html=True)
        monto = st.number_input("Monto (MXN)", min_value=0.0, step=10.0, format="%.2f")
        
        st.markdown('<hr class="divider">', unsafe_allow_html=True)
        st.markdown('<span class="step-badge">4</span> **Comprobante**', unsafe_allow_html=True)
        comprobante = st.file_uploader(
            "Sube foto o PDF del comprobante",
            type=["jpg", "jpeg", "png", "pdf"],
            help="Foto clara del comprobante físico, o PDF si es digital."
        )
        
        notas = st.text_area("Notas (opcional)", placeholder="Cualquier observación adicional...", height=80)
        
        st.markdown('<hr class="divider">', unsafe_allow_html=True)
        submitted = st.form_submit_button("✅ Registrar Entrega")

    if submitted:
        errores = []
        if not rider:
            errores.append("Selecciona un rider.")
        if not agencia:
            errores.append("Selecciona una agencia.")
        if not recibido_por:
            errores.append("Indica quién recibió.")
        if monto <= 0:
            errores.append("El monto debe ser mayor a $0.")
        if not comprobante:
            errores.append("Debes subir el comprobante.")
        
        if errores:
            for e in errores:
                st.markdown(f'<div class="error-box">⚠️ {e}</div>', unsafe_allow_html=True)
        else:
            with st.spinner("Guardando registro..."):
                try:
                    # Hoja semanal
                    ws, _ = ensure_week_sheet(spreadsheet)
                    folio = get_next_folio(ws)
                    fecha_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
                    
                    # Subir comprobante a ImgBB
                    file_bytes = comprobante.read()
                    filename = f"{folio}_{rider.replace(' ', '_')}"
                    comp_url = upload_comprobante_imgbb(file_bytes, filename)
                    
                    # Registrar en Sheets
                    append_registro(ws, spreadsheet, folio, fecha_hora, rider, agencia, monto, recibido_por, notas, comp_url)
                    
                    st.session_state.registro_exitoso = folio
                    st.rerun()

                except Exception as e:
                    import traceback
                    st.error(f"Error: {str(e)}")
                    st.code(traceback.format_exc())


# ─────────────────────────────────────────────
# VISTA: ADMIN
# ─────────────────────────────────────────────
elif st.session_state.vista == "admin":

    if not st.session_state.admin_logged:
        st.markdown("### 🔐 Acceso Administrador")
        pwd = st.text_input("Contraseña", type="password")
        if st.button("Entrar"):
            if pwd == ADMIN_PASSWORD:
                st.session_state.admin_logged = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

    # ADMIN PANEL
    st.markdown("## ⚙️ Panel de Administración")

    try:
        client, creds = get_google_client()
        spreadsheet_id = st.secrets["spreadsheet_id"]
        folder_id = st.secrets["drive_folder_id"]
        spreadsheet = get_or_create_sheet(client, spreadsheet_id)
        riders, agencias = load_config(spreadsheet)
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        st.stop()

    tab1, tab2, tab3 = st.tabs(["👥 Riders", "🏢 Agencias y Encargados", "📊 Reporte Semanal"])

    # ── TAB 1: Riders ──
    with tab1:
        st.markdown('<div class="admin-section">', unsafe_allow_html=True)
        st.markdown("**Lista de Riders activos**")
        
        riders_text = st.text_area(
            "Un rider por línea",
            value="\n".join(riders),
            height=220
        )
        
        if st.button("💾 Guardar Riders"):
            new_riders = [r.strip() for r in riders_text.split("\n") if r.strip()]
            save_config(spreadsheet, new_riders, agencias)
            st.success(f"✅ {len(new_riders)} riders guardados.")
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── TAB 2: Agencias ──
    with tab2:
        st.markdown('<div class="admin-section">', unsafe_allow_html=True)
        st.markdown("**Agencias y sus encargados**")
        
        new_agencias = {}
        for agencia_name in list(agencias.keys()):
            st.markdown(f"**🏢 {agencia_name}**")
            enc_text = st.text_area(
                f"Encargados de {agencia_name} (uno por línea)",
                value="\n".join(agencias[agencia_name]),
                height=100,
                key=f"enc_{agencia_name}"
            )
            new_agencias[agencia_name] = [e.strip() for e in enc_text.split("\n") if e.strip()]
            st.markdown("---")
        
        # Agregar agencia nueva
        st.markdown("**➕ Agregar nueva agencia**")
        nueva_agencia = st.text_input("Nombre de la nueva agencia", key="nueva_ag")
        if st.button("Agregar agencia"):
            if nueva_agencia and nueva_agencia not in new_agencias:
                new_agencias[nueva_agencia] = []
                save_config(spreadsheet, riders, new_agencias)
                st.success(f"✅ Agencia '{nueva_agencia}' agregada.")
                st.rerun()
        
        if st.button("💾 Guardar Encargados"):
            save_config(spreadsheet, riders, new_agencias)
            st.success("✅ Encargados guardados.")
            st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)

    # ── TAB 3: Reporte Semanal ──
    with tab3:
        st.markdown('<div class="admin-section">', unsafe_allow_html=True)
        
        # Listar hojas disponibles
        sheet_names = [ws.title for ws in spreadsheet.worksheets() if ws.title != "Config"]
        
        if not sheet_names:
            st.info("No hay registros aún.")
        else:
            semana_sel = st.selectbox("Selecciona semana", options=sheet_names[::-1])
            
            ws = spreadsheet.worksheet(semana_sel)
            records = ws.get_all_values()
            
            if len(records) < 2:
                st.info("Sin registros esta semana.")
            else:
                # Filtrar solo filas de datos (folio empieza con NIRO-)
                data_rows = [r for r in records[1:] if r[0].startswith("NIRO-")]
                
                if not data_rows:
                    st.info("Sin registros válidos.")
                else:
                    df = pd.DataFrame(data_rows, columns=records[0][:len(data_rows[0])])
                    df["Monto (MXN)"] = pd.to_numeric(df["Monto (MXN)"], errors="coerce").fillna(0)
                    
                    total = df["Monto (MXN)"].sum()
                    
                    # KPIs
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Total Semana", f"${total:,.2f}")
                    k2.metric("Registros", len(df))
                    k3.metric("Riders activos", df["Rider"].nunique())
                    
                    st.markdown("---")
                    
                    # Tabla
                    st.markdown("**📋 Registros**")
                    st.dataframe(
                        df[["Folio", "Fecha y Hora", "Rider", "Agencia", "Monto (MXN)", "Recibido por", "Notas"]],
                        use_container_width=True,
                        hide_index=True
                    )
                    
                    st.markdown("---")
                    
                    # Totales por Rider
                    st.markdown("**💰 Total por Rider**")
                    por_rider = df.groupby("Rider")["Monto (MXN)"].sum().reset_index()
                    por_rider = por_rider.sort_values("Monto (MXN)", ascending=False)
                    por_rider["Monto (MXN)"] = por_rider["Monto (MXN)"].apply(lambda x: f"${x:,.2f}")
                    st.dataframe(por_rider, use_container_width=True, hide_index=True)
                    
                    st.markdown("---")
                    
                    # PIE CHART
                    st.markdown("**🥧 Distribución por Rider**")
                    pie_data = df.groupby("Rider")["Monto (MXN)"].sum()
                    st.pyplot(generate_pie_chart(pie_data))
                    
                    st.markdown("---")
                    
                    # Botón para actualizar resumen en Sheets
                    if st.button("📤 Actualizar resumen en Google Sheets"):
                        with st.spinner("Actualizando..."):
                            update_weekly_report(ws, spreadsheet_id, client, creds)
                            st.success("✅ Resumen actualizado en Google Sheets.")
        
        st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# PIE CHART helper (fuera del flujo principal)
# ─────────────────────────────────────────────
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

def generate_pie_chart(pie_data):
    colors = ["#E8450A", "#FF7043", "#FF8A65", "#FFAB91", "#FFCCBC",
              "#BF360C", "#D84315", "#F4511E"]
    
    fig, ax = plt.subplots(figsize=(7, 5), facecolor="#1a1a1a")
    wedges, texts, autotexts = ax.pie(
        pie_data.values,
        labels=pie_data.index,
        autopct="%1.1f%%",
        colors=colors[:len(pie_data)],
        startangle=140,
        wedgeprops={"edgecolor": "#0f0f0f", "linewidth": 2}
    )
    for t in texts:
        t.set_color("#cccccc")
        t.set_fontsize(9)
    for at in autotexts:
        at.set_color("white")
        at.set_fontweight("bold")
        at.set_fontsize(9)
    
    ax.set_title("Distribución de Efectivo por Rider", color="white", fontsize=12, fontweight="bold", pad=15)
    fig.patch.set_facecolor("#1a1a1a")
    plt.tight_layout()
    return fig
