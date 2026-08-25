import base64
import hashlib
import io
import math
import sqlite3
import uuid
from datetime import date

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from streamlit_geolocation import streamlit_geolocation

# MÓDULO SUPABASE
from supabase import Client, create_client

# IMPORTAÇÃO DOS MÓDULOS CUSTOMIZADOS
from ui_components import aplicar_estilo_customizado, render_kpi_card

# ==============================================================================
# INICIALIZAÇÃO DA APLICAÇÃO E SESSION STATE (EVITA KEYERROR)
# ==============================================================================

st.set_page_config(page_title="Central de Controle - Sondagem", layout="wide")

# Garantir que todas as chaves de sessão existam antes de qualquer leitura
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario" not in st.session_state:
    st.session_state["usuario"] = ""
if "perfil" not in st.session_state:
    st.session_state["perfil"] = ""
if "sonda_id" not in st.session_state:
    st.session_state["sonda_id"] = None

# ==============================================================================
# CONEXÃO E CONFIGURAÇÃO DO SUPABASE
# ==============================================================================

SUPABASE_URL = st.secrets.get("SUPABASE_URL", "https://seu-projeto.supabase.co")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "sua-chave-aqui")


@st.cache_resource
def get_supabase_client() -> Client:
    """Inicializa e retorna a conexão com o Supabase."""
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def upload_foto_supabase(file_buffer, nome_arquivo):
    """Envia uma imagem para o Supabase Storage e retorna a URL pública."""
    try:
        supabase = get_supabase_client()
        bucket_name = "boletins-fotos"

        # Gera nome único para evitar sobrescrever arquivos
        ext = nome_arquivo.split(".")[-1]
        caminho_storage = f"boletins/{uuid.uuid4().hex}.{ext}"

        # Upload para o Storage
        supabase.storage.from_(bucket_name).upload(
            caminho_storage,
            file_buffer.getvalue(),
            file_options={"content-type": f"image/{ext}"},
        )

        # Retorna URL pública do arquivo
        public_url = supabase.storage.from_(bucket_name).get_public_url(
            caminho_storage
        )
        return public_url
    except Exception as e:
        st.error(f"Erro ao enviar imagem para o Supabase Storage: {e}")
        return None


def salvar_boletim_supabase(dados_boletim):
    """Insere o registro do boletim geológico no banco PostgreSQL do Supabase."""
    try:
        supabase = get_supabase_client()
        response = (
            supabase.table("boletim_geologico").insert(dados_boletim).execute()
        )
        return response
    except Exception as e:
        st.error(f"Erro ao salvar boletim no Supabase: {e}")
        return None


# ==============================================================================
# AUTENTICAÇÃO E GESTÃO DE USUÁRIOS (RBAC) & BANCO LOCAL SQLITE
# ==============================================================================


def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()


def get_connection():
    return sqlite3.connect("central_sondagem.db")


def criar_tabela_usuarios():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sondas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE NOT NULL,
            equipe TEXT NOT NULL,
            projeto TEXT NOT NULL,
            status TEXT CHECK(status IN ('Operando', 'Parada', 'Manutenção')) DEFAULT 'Operando'
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            senha TEXT NOT NULL,
            perfil TEXT CHECK(perfil IN ('Admin', 'Geólogo', 'Operador')) NOT NULL,
            sonda_id INTEGER,
            FOREIGN KEY(sonda_id) REFERENCES sondas(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS furos (
            id TEXT PRIMARY KEY,
            sonda_id INTEGER,
            coord_e REAL,
            coord_n REAL,
            cota REAL,
            prof_planejada REAL,
            prof_executada REAL DEFAULT 0,
            situacao TEXT CHECK(situacao IN ('Planejado', 'Em Andamento', 'Concluído', 'Cancelado')) DEFAULT 'Planejado',
            FOREIGN KEY(sonda_id) REFERENCES sondas(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS producao_diaria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data DATE NOT NULL,
            sonda_id INTEGER,
            furo_id TEXT,
            prof_inicial REAL NOT NULL,
            prof_final REAL NOT NULL,
            horas_trabalhadas REAL NOT NULL,
            horas_paradas REAL DEFAULT 0,
            motivo_parada TEXT,
            FOREIGN KEY(sonda_id) REFERENCES sondas(id),
            FOREIGN KEY(furo_id) REFERENCES furos(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS boletim_geologico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            furo_id TEXT NOT NULL,
            de_m REAL NOT NULL,
            ate_m REAL NOT NULL,
            recuperacao_m REAL NOT NULL,
            rqd_m REAL NOT NULL,
            litologia TEXT NOT NULL,
            descricao_geologica TEXT,
            n_amostra TEXT,
            observacoes TEXT,
            foto_url TEXT,
            FOREIGN KEY(furo_id) REFERENCES furos(id)
        )
    """)

    cursor.execute("SELECT COUNT(*) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            "INSERT INTO usuarios (usuario, senha, perfil, sonda_id) VALUES (?, ?, ?, NULL)",
            ("admin", hash_senha("admin123"), "Admin"),
        )

    conn.commit()
    conn.close()


def verificar_login(usuario, senha):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT perfil, sonda_id FROM usuarios WHERE usuario = ? AND senha = ?",
        (usuario, hash_senha(senha)),
    )
    resultado = cursor.fetchone()
    conn.close()
    return resultado if resultado else None


def tela_login():
    criar_tabela_usuarios()

    if not st.session_state.get("autenticado", False):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.subheader("🔐 Acesso ao Sistema")
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")

            if st.button("Entrar", type="primary", use_container_width=True):
                login_info = verificar_login(usuario, senha)
                if login_info:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = usuario
                    st.session_state["perfil"] = login_info[0]
                    st.session_state["sonda_id"] = login_info[1]
                    st.success("Login realizado com sucesso!")
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
        return False
    return True


def botao_logout():
    st.sidebar.markdown(
        f"👤 **{st.session_state.get('usuario', '')}** ({st.session_state.get('perfil', '')})"
    )
    if st.sidebar.button("Sair / Logout"):
        st.session_state["autenticado"] = False
        st.session_state["usuario"] = ""
        st.session_state["perfil"] = ""
        st.session_state["sonda_id"] = None
        st.rerun()


# ==============================================================================
# UTILITÁRIOS E EXPORTAÇÃO EXCEL
# ==============================================================================


def latlon_to_utm(lat, lon):
    if lat is None or lon is None or (lat == 0.0 and lon == 0.0):
        return 0.0, 0.0, 0
    a = 6378137.0
    f = 1 / 298.257223563
    b = a * (1 - f)
    e_sq = (a**2 - b**2) / (a**2)
    k0 = 0.9996
    zone_number = int((lon + 180) / 6) + 1
    lon0 = (zone_number - 1) * 6 - 180 + 3
    lat_rad = math.radians(lat)
    lon_rad = math.radians(lon)
    lon0_rad = math.radians(lon0)
    N = a / math.sqrt(1 - e_sq * math.sin(lat_rad) ** 2)
    T = math.tan(lat_rad) ** 2
    C = (e_sq / (1 - e_sq)) * math.cos(lat_rad) ** 2
    A = (lon_rad - lon0_rad) * math.cos(lat_rad)
    M = a * (
        (1 - e_sq / 4 - 3 * e_sq**2 / 64 - 5 * e_sq**3 / 256) * lat_rad
        - (3 * e_sq / 8 + 3 * e_sq**2 / 32 + 45 * e_sq**3 / 1024)
        * math.sin(2 * lat_rad)
        + (15 * e_sq**2 / 256 + 45 * e_sq**3 / 1024) * math.sin(4 * lat_rad)
        - (35 * e_sq**3 / 3072) * math.sin(6 * lat_rad)
    )
    easting = (
        k0
        * N
        * (
            A
            + (1 - T + C) * A**3 / 6
            + (5 - 18 * T + T**2 + 72 * C - 58 * (e_sq / (1 - e_sq)))
            * A**5
            / 120
        )
        + 500000.0
    )
    northing = k0 * (
        M
        + N
        * math.tan(lat_rad)
        * (
            A**2 / 2
            + (5 - T + 9 * C + 4 * C**2) * A**4 / 24
            + (61 - 58 * T + T**2 + 600 * C - 330 * (e_sq / (1 - e_sq)))
            * A**6
            / 720
        )
    )
    if lat < 0:
        northing += 10000000.0
    return round(easting, 2), round(northing, 2), zone_number


def add_bg_from_local(image_file):
    try:
        with open(image_file, "rb") as image:
            encoded_string = base64.b64encode(image.read())
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/png;base64,{encoded_string.decode()}");
                background-size: cover; background-position: center;
                background-repeat: no-repeat; background-attachment: fixed;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )
    except FileNotFoundError:
        pass


def gerar_dashboard_excel_completo(perfil_usuario, user_sonda_id):
    output = io.BytesIO()
    wb = openpyxl.Workbook()

    cor_cabecalho = PatternFill(
        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
    )
    cor_titulo = PatternFill(
        start_color="1B365D", end_color="1B365D", fill_type="solid"
    )
    cor_card = PatternFill(
        start_color="F2F4F7", end_color="F2F4F7", fill_type="solid"
    )
    cor_zebra = PatternFill(
        start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
    )

    fonte_titulo = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fonte_card_num = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    fonte_card_lbl = Font(name="Calibri", size=9, italic=True, color="595959")
    fonte_dados = Font(name="Calibri", size=10)

    borda_fina = Side(border_style="thin", color="D9D9D9")
    borda_caixa = Border(
        left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina
    )

    conn = get_connection()

    if perfil_usuario == "Admin":
        df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
        df_prod = pd.read_sql_query(
            "SELECT p.*, s.codigo as sonda_codigo FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id",
            conn,
        )
        df_geo = pd.read_sql_query(
            """
            SELECT id, furo_id, de_m, ate_m, (ate_m - de_m) as avanco_m, recuperacao_m, 
                   ROUND((recuperacao_m / NULLIF(ate_m - de_m, 0)) * 100, 1) as recuperacao_pct,
                   rqd_m, ROUND((rqd_m / NULLIF(ate_m - de_m, 0)) * 100, 1) as rqd_pct, litologia, n_amostra, descricao_geologica, observacoes, foto_url
            FROM boletim_geologico ORDER BY furo_id, de_m ASC
        """,
            conn,
        )
    else:
        df_sondas = pd.read_sql_query(
            "SELECT * FROM sondas WHERE id = ?", conn, params=(user_sonda_id,)
        )
        df_prod = pd.read_sql_query(
            "SELECT p.*, s.codigo as sonda_codigo FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id WHERE p.sonda_id = ?",
            conn,
            params=(user_sonda_id,),
        )
        df_geo = pd.read_sql_query(
            """
            SELECT bg.id, bg.furo_id, bg.de_m, bg.ate_m, (bg.ate_m - bg.de_m) as avanco_m, bg.recuperacao_m, 
                   ROUND((bg.recuperacao_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as recuperacao_pct,
                   bg.rqd_m, ROUND((bg.rqd_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as rqd_pct, bg.litologia, bg.n_amostra, bg.descricao_geologica, bg.observacoes, bg.foto_url
            FROM boletim_geologico bg JOIN furos f ON bg.furo_id = f.id WHERE f.sonda_id = ? ORDER BY bg.furo_id, bg.de_m ASC
        """,
            conn,
            params=(user_sonda_id,),
        )
    conn.close()

    ws_dash = wb.active
    ws_dash.title = "📌 Dashboard Executivo"
    ws_dash.views.sheetView[0].showGridLines = True
    ws_dash.merge_cells("A1:H1")
    ws_dash["A1"] = "DASHBOARD EXECUTIVO — CONTROLE INTEGRADO DE SONDAGEM"
    ws_dash["A1"].font = fonte_titulo
    ws_dash["A1"].fill = cor_titulo
    ws_dash["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[1].height = 40

    total_sondas = len(df_sondas)
    sondas_op = len(df_sondas[df_sondas["status"] == "Operando"])
    total_metros = (
        (df_prod["prof_final"] - df_prod["prof_inicial"]).sum()
        if not df_prod.empty
        else 0.0
    )
    total_hrs_trab = (
        df_prod["horas_trabalhadas"].sum() if not df_prod.empty else 0.0
    )
    total_hrs_par = (
        df_prod["horas_paradas"].sum() if not df_prod.empty else 0.0
    )
    eficiencia = (
        (total_hrs_trab / (total_hrs_trab + total_hrs_par) * 100)
        if (total_hrs_trab + total_hrs_par) > 0
        else 0.0
    )

    kpis = [
        ("Sondas Operando", f"{sondas_op}/{total_sondas}", "A3:B3", "A4:B4", "A3:B4"),
        ("Metragem Total", f"{total_metros:.1f} m", "D3:E3", "D4:E4", "D3:E4"),
        ("Eficiência Geral", f"{eficiencia:.1f}%", "G3:H3", "G4:H4", "G3:H4"),
    ]

    for lbl, val, r_lbl, r_val, r_box in kpis:
        ws_dash.merge_cells(r_lbl)
        ws_dash[r_lbl.split(":")[0]] = lbl
        ws_dash[r_lbl.split(":")[0]].font = fonte_card_lbl
        ws_dash[r_lbl.split(":")[0]].alignment = Alignment(horizontal="center")
        ws_dash.merge_cells(r_val)
        ws_dash[r_val.split(":")[0]] = val
        ws_dash[r_val.split(":")[0]].font = fonte_card_num
        ws_dash[r_val.split(":")[0]].alignment = Alignment(horizontal="center")
        for row in ws_dash[r_box]:
            for cell in row:
                cell.fill = cor_card
                cell.border = borda_caixa

    # Sondas Sheet
    ws_sondas = wb.create_sheet(title="🚜 Sondas & Equipes")
    ws_sondas.views.sheetView[0].showGridLines = True
    headers_sondas = [
        "ID",
        "Código Sonda",
        "Equipe Responsável",
        "Projeto / Frente",
        "Status Atual",
    ]
    for c_idx, h in enumerate(headers_sondas, 1):
        cell = ws_sondas.cell(row=1, column=c_idx, value=h)
        cell.font = fonte_cabecalho
        cell.fill = cor_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_caixa
    for r_idx, row in df_sondas.iterrows():
        curr_row = r_idx + 2
        values = [
            row["id"],
            row["codigo"],
            row["equipe"],
            row["projeto"],
            row["status"],
        ]
        for c_idx, val in enumerate(values, 1):
            cell = ws_sondas.cell(row=curr_row, column=c_idx, value=val)
            cell.font = fonte_dados
            cell.border = borda_caixa
            if r_idx % 2 == 0:
                cell.fill = cor_zebra

    # Produção Sheet
    ws_prod = wb.create_sheet(title="📝 Apontamento Diário")
    ws_prod.views.sheetView[0].showGridLines = True
    headers_prod = [
        "Data",
        "Sonda",
        "Furo",
        "Prof. Inicial (m)",
        "Prof. Final (m)",
        "Avanço (m)",
        "Horas Trab.",
        "Horas Paradas",
        "Motivo Parada",
    ]
    for c_idx, h in enumerate(headers_prod, 1):
        cell = ws_prod.cell(row=1, column=c_idx, value=h)
        cell.font = fonte_cabecalho
        cell.fill = cor_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_caixa
    for r_idx, row in df_prod.iterrows():
        curr_row = r_idx + 2
        avanco = row["prof_final"] - row["prof_inicial"]
        values = [
            row["data"],
            row["sonda_codigo"],
            row["furo_id"],
            row["prof_inicial"],
            row["prof_final"],
            avanco,
            row["horas_trabalhadas"],
            row["horas_paradas"],
            row["motivo_parada"],
        ]
        for c_idx, val in enumerate(values, 1):
            cell = ws_prod.cell(row=curr_row, column=c_idx, value=val)
            cell.font = fonte_dados
            cell.border = borda_caixa
            if r_idx % 2 == 0:
                cell.fill = cor_zebra

    # Boletim Sheet
    ws_geo = wb.create_sheet(title="⛏️ Boletim Geológico")
    ws_geo.views.sheetView[0].showGridLines = True
    headers_geo = [
        "Furo",
        "De (m)",
        "Até (m)",
        "Avanço (m)",
        "Rec. (m)",
        "Rec. (%)",
        "RQD (m)",
        "RQD (%)",
        "Litologia",
        "Amostra",
        "Descrição Geológica",
        "Observações",
        "URL Foto Supabase",
    ]
    for c_idx, h in enumerate(headers_geo, 1):
        cell = ws_geo.cell(row=1, column=c_idx, value=h)
        cell.font = fonte_cabecalho
        cell.fill = cor_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_caixa
    for r_idx, row in df_geo.iterrows():
        curr_row = r_idx + 2
        values = [
            row["furo_id"],
            row["de_m"],
            row["ate_m"],
            row["avanco_m"],
            row["recuperacao_m"],
            row["recuperacao_pct"],
            row["rqd_m"],
            row["rqd_pct"],
            row["litologia"],
            row["n_amostra"],
            row["descricao_geologica"],
            row["observacoes"],
            row.get("foto_url", ""),
        ]
        for c_idx, val in enumerate(values, 1):
            cell = ws_geo.cell(row=curr_row, column=c_idx, value=val)
            cell.font = fonte_dados
            cell.border = borda_caixa
            if r_idx % 2 == 0:
                cell.fill = cor_zebra

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# CARREGAMENTO DE ESTILOS E AUTENTICAÇÃO
# ==============================================================================

add_bg_from_local("logo_empresa.png")
aplicar_estilo_customizado()

if not tela_login():
    st.stop()

perfil_atual = st.session_state.get("perfil", "")
sonda_id_atual = st.session_state.get("sonda_id", None)

# ==============================================================================
# NAVEGAÇÃO STREAMLIT & SIDEBAR
# ==============================================================================

st.sidebar.title("🛠️ CENTRAL DE CONTROLE")
botao_logout()
st.sidebar.markdown("---")

opcoes_menu = [
    "📊 Dashboard Geral",
    "🚜 Cadastro de Sondas",
    "📝 Apontamento Diário",
    "📍 Controle de Furos",
    "⛏️ Boletim Geológico",
]

if perfil_atual == "Admin":
    opcoes_menu.append("👥 Gestão de Usuários")

opcao = st.sidebar.radio("Navegação", opcoes_menu)

excel_mestre = gerar_dashboard_excel_completo(perfil_atual, sonda_id_atual)
st.sidebar.markdown("---")
st.sidebar.download_button(
    label="📊 Exportar Planilha Mestra",
    data=excel_mestre,
    file_name=f"Dashboard_Sondagem_{date.today()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

# ------------------------------------------------------------------------------
# 1. DASHBOARD GERAL
# ------------------------------------------------------------------------------
if opcao == "📊 Dashboard Geral":
    st.title("📊 Painel Geral de Operações")
    st.caption("Visão macro da produção e disponibilidade de equipamentos.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
        df_prod = pd.read_sql_query(
            "SELECT p.*, s.codigo as sonda_codigo FROM producao_diaria p JOIN sondas s ON p.sonda_id = s.id",
            conn,
        )
    else:
        df_sondas = pd.read_sql_query(
            "SELECT * FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,)
        )
        df_prod = pd.read_sql_query(
            "SELECT p.*, s.codigo as sonda_codigo FROM producao_diaria p JOIN sondas s ON p.sonda_id = s.id WHERE p.sonda_id = ?",
            conn,
            params=(sonda_id_atual,),
        )
    conn.close()

    sondas_total = len(df_sondas)
    sondas_op = len(df_sondas[df_sondas["status"] == "Operando"])

    if not df_prod.empty:
        metros_hoje = (
            df_prod[df_prod["data"] == str(date.today())]["prof_final"].sum()
            - df_prod[df_prod["data"] == str(date.today())]["prof_inicial"].sum()
        )
        metros_acumulados = (df_prod["prof_final"] - df_prod["prof_inicial"]).sum()
        hrs_trab = df_prod["horas_trabalhadas"].sum()
        hrs_par = df_prod["horas_paradas"].sum()
        eficiencia = (
            (hrs_trab / (hrs_trab + hrs_par) * 100) if (hrs_trab + hrs_par) > 0 else 0
        )
    else:
        metros_hoje, metros_acumulados, eficiencia = 0.0, 0.0, 0.0

    media_por_sonda = metros_hoje / sondas_op if sondas_op > 0 else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        render_kpi_card("Sondas Operando", f"{sondas_op}/{sondas_total}")
    with c2:
        render_kpi_card("Produção Hoje", f"{metros_hoje:.1f} m")
    with c3:
        render_kpi_card("Total Acumulado", f"{metros_acumulados:.1f} m")
    with c4:
        render_kpi_card("Média / Sonda", f"{media_por_sonda:.1f} m")
    with c5:
        render_kpi_card("Eficiência", f"{eficiencia:.0f}%")

    st.markdown("---")
    st.subheader("Status das Sondas em Campo")

    if not df_sondas.empty:
        cols = st.columns(min(max(sondas_total, 1), 4))
        status_emojis = {"Operando": "🟢", "Parada": "🟡", "Manutenção": "🔴"}
        for i, (_, sonda) in enumerate(df_sondas.iterrows()):
            with cols[i % 4]:
                with st.container(border=True):
                    emoji = status_emojis.get(sonda["status"], "⚪")
                    st.markdown(f"### {emoji} {sonda['codigo']}")
                    st.write(f"**Equipe:** {sonda['equipe']}")
                    st.write(f"**Projeto:** {sonda['projeto']}")
                    st.write(f"**Status:** {sonda['status']}")
    else:
        st.info("Nenhuma sonda vinculada a este perfil.")

# ------------------------------------------------------------------------------
# 2. CADASTRO DE SONDAS
# ------------------------------------------------------------------------------
elif opcao == "🚜 Cadastro de Sondas":
    st.title("🚜 Gestão de Sondas")
    st.caption("Controle e atualização do parque de equipamentos.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
    else:
        df_sondas = pd.read_sql_query(
            "SELECT * FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,)
        )
    conn.close()

    if perfil_atual == "Admin":
        tab_lista, tab_novo, tab_editar = st.tabs(
            ["📋 Sondas Cadastradas", "➕ Nova Sonda", "✏️ Editar Sonda"]
        )
    else:
        tab_lista = st.container()

    with tab_lista:
        if not df_sondas.empty:
            st.dataframe(df_sondas, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma sonda encontrada.")

    if perfil_atual == "Admin":
        with tab_novo:
            with st.container(border=True):
                with st.form("form_sonda", clear_on_submit=True):
                    st.subheader("Cadastrar Nova Sonda")
                    c1, c2 = st.columns(2)
                    codigo = c1.text_input("Identificação (ex: SD-01)")
                    equipe = c2.text_input("Equipe Responsável")

                    c3, c4 = st.columns(2)
                    projeto = c3.text_input("Frente / Projeto")
                    status = c4.selectbox(
                        "Status Atual", ["Operando", "Parada", "Manutenção"]
                    )

                    btn_salvar = st.form_submit_button(
                        "Cadastrar Sonda", type="primary", use_container_width=True
                    )

                    if btn_salvar and codigo and equipe and projeto:
                        conn = get_connection()
                        cursor = conn.cursor()
                        try:
                            cursor.execute(
                                "INSERT INTO sondas (codigo, equipe, projeto, status) VALUES (?, ?, ?, ?)",
                                (codigo, equipe, projeto, status),
                            )
                            conn.commit()
                            st.success(f"Sonda '{codigo}' cadastrada!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Esta identificação já existe.")
                        finally:
                            conn.close()

        with tab_editar:
            if not df_sondas.empty:
                with st.container(border=True):
                    with st.form("form_edit_sonda"):
                        st.subheader("Editar Dados da Sonda")
                        sonda_para_editar = st.selectbox(
                            "Selecione a Sonda", df_sondas["codigo"].tolist()
                        )
                        dados_sonda = df_sondas[
                            df_sondas["codigo"] == sonda_para_editar
                        ].iloc[0]

                        c1, c2 = st.columns(2)
                        novo_codigo = c1.text_input(
                            "Código Oficial", value=dados_sonda["codigo"]
                        )
                        nova_equipe = c2.text_input(
                            "Equipe", value=dados_sonda["equipe"]
                        )

                        c3, c4 = st.columns(2)
                        novo_projeto = c3.text_input(
                            "Projeto", value=dados_sonda["projeto"]
                        )
                        novo_status = c4.selectbox(
                            "Status",
                            ["Operando", "Parada", "Manutenção"],
                            index=["Operando", "Parada", "Manutenção"].index(
                                dados_sonda["status"]
                            ),
                        )

                        btn_atualizar = st.form_submit_button(
                            "Atualizar Cadastro",
                            type="primary",
                            use_container_width=True,
                        )

                        if btn_atualizar:
                            conn = get_connection()
                            cursor = conn.cursor()
                            try:
                                cursor.execute(
                                    "UPDATE sondas SET codigo = ?, equipe = ?, projeto = ?, status = ? WHERE id = ?",
                                    (
                                        novo_codigo,
                                        nova_equipe,
                                        novo_projeto,
                                        novo_status,
                                        int(dados_sonda["id"]),
                                    ),
                                )
                                conn.commit()
                                st.success(f"Sonda '{novo_codigo}' atualizada!")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("O novo código já pertence a outra sonda.")
                            finally:
                                conn.close()

# ------------------------------------------------------------------------------
# 3. APONTAMENTO DIÁRIO
# ------------------------------------------------------------------------------
elif opcao == "📝 Apontamento Diário":
    st.title("📝 Apontamento Diário de Produção")
    st.caption("Registro de avanço físico e tempos de paralisação.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
        df_furos = pd.read_sql_query("SELECT id, sonda_id FROM furos", conn)
        df_prod_full = pd.read_sql_query(
            """
            SELECT p.id, p.data, s.codigo as sonda_codigo, p.furo_id, p.prof_inicial, p.prof_final, 
                   (p.prof_final - p.prof_inicial) as avanco, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada
            FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id
            ORDER BY p.data DESC, p.id DESC
        """,
            conn,
        )
    else:
        df_sondas = pd.read_sql_query(
            "SELECT id, codigo FROM sondas WHERE id = ?",
            conn,
            params=(sonda_id_atual,),
        )
        df_furos = pd.read_sql_query(
            "SELECT id, sonda_id FROM furos WHERE sonda_id = ?",
            conn,
            params=(sonda_id_atual,),
        )
        df_prod_full = pd.read_sql_query(
            """
            SELECT p.id, p.data, s.codigo as sonda_codigo, p.furo_id, p.prof_inicial, p.prof_final, 
                   (p.prof_final - p.prof_inicial) as avanco, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada
            FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id
            WHERE p.sonda_id = ? ORDER BY p.data DESC, p.id DESC
        """,
            conn,
            params=(sonda_id_atual,),
        )
    conn.close()

    tab_historico, tab_novo, tab_excluir = st.tabs(
        ["📋 Histórico", "➕ Novo Apontamento", "🗑️ Excluir Registro"]
    )

    with tab_historico:
        if not df_prod_full.empty:
            st.dataframe(df_prod_full, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum apontamento registrado.")

    with tab_novo:
        if not df_sondas.empty and not df_furos.empty:
            with st.container(border=True):
                with st.form("form_producao", clear_on_submit=True):
                    st.subheader("Registrar Avanço Diário")
                    c1, c2, c3 = st.columns(3)
                    data_reg = c1.date_input("Data do Avanço", value=date.today())
                    
                    sonda_sel = c2.selectbox(
                        "Sonda",
                        df_sondas["codigo"].tolist()
                    )
                    sonda_id_sel = df_sondas[df_sondas["codigo"] == sonda_sel]["id"].values[0]

                    furos_filtrados = df_furos[df_furos["sonda_id"] == sonda_id_sel]["id"].tolist()
                    furo_sel = c3.selectbox("Furo", furos_filtrados if furos_filtrados else ["Nenhum furo cadastrado"])

                    c4, c5 = st.columns(2)
                    prof_ini = c4.number_input("Profundidade Inicial (m)", min_value=0.0, step=0.5)
                    prof_fin = c5.number_input("Profundidade Final (m)", min_value=0.0, step=0.5)

                    c6, c7, c8 = st.columns(3)
                    hrs_trab = c6.number_input("Horas Trabalhadas", min_value=0.0, max_value=24.0, step=0.5)
                    hrs_par = c7.number_input("Horas Paradas", min_value=0.0, max_value=24.0, step=0.5)
                    motivo_parada = c8.text_input("Motivo da Parada")

                    btn_salvar_prod = st.form_submit_button("Salvar Apontamento", type="primary", use_container_width=True)

                    if btn_salvar_prod:
                        if prof_fin < prof_ini:
                            st.error("A profundidade final deve ser maior ou igual à inicial.")
                        elif furo_sel == "Nenhum furo cadastrado":
                            st.error("Selecione um furo válido.")
                        else:
                            conn = get_connection()
                            cursor = conn.cursor()
                            cursor.execute(
                                """
                                INSERT INTO producao_diaria (data, sonda_id, furo_id, prof_inicial, prof_final, horas_trabalhadas, horas_paradas, motivo_parada)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                                """,
                                (data_reg, int(sonda_id_sel), furo_sel, prof_ini, prof_fin, hrs_trab, hrs_par, motivo_parada)
                            )
                            cursor.execute(
                                "UPDATE furos SET prof_executada = ?, situacao = 'Em Andamento' WHERE id = ?",
                                (prof_fin, furo_sel)
                            )
                            conn.commit()
                            conn.close()
                            st.success("Apontamento registrado com sucesso!")
                            st.rerun()
        else:
            st.warning("É necessário ter pelo menos uma sonda e um furo cadastrados para registrar produções.")

    with tab_excluir:
        if perfil_atual in ["Admin", "Geólogo"]:
            if not df_prod_full.empty:
                reg_id = st.selectbox("Selecione o ID do registro para excluir", df_prod_full["id"].tolist())
                if st.button("Confirmar Exclusão", type="primary"):
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM producao_diaria WHERE id = ?", (reg_id,))
                    conn.commit()
                    conn.close()
                    st.success(f"Registro ID {reg_id} removido!")
                    st.rerun()
            else:
                st.info("Nenhum registro para excluir.")
        else:
            st.error("Perfil de Operador não tem permissão para excluir registros.")

# ------------------------------------------------------------------------------
# 4. CONTROLE DE FUROS
# ------------------------------------------------------------------------------
elif opcao == "📍 Controle de Furos":
    st.title("📍 Controle de Furos de Sondagem")
    st.caption("Planejamento, coordenadas geográficas e status de perfuração.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_furos = pd.read_sql_query("SELECT f.*, s.codigo as sonda_codigo FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id", conn)
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
    else:
        df_furos = pd.read_sql_query("SELECT f.*, s.codigo as sonda_codigo FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id WHERE f.sonda_id = ?", conn, params=(sonda_id_atual,))
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,))
    conn.close()

    tab_furos_lista, tab_furo_novo = st.tabs(["📋 Furos Cadastrados", "➕ Novo Furo"])

    with tab_furos_lista:
        if not df_furos.empty:
            st.dataframe(df_furos, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum furo cadastrado.")

    with tab_furo_novo:
        if not df_sondas.empty:
            with st.container(border=True):
                st.subheader("Cadastrar Novo Furo")
                
                st.markdown("##### 🛰️ Obter Coordenadas via GPS")
                location = streamlit_geolocation()
                lat_gps = location.get("latitude") if location else None
                lon_gps = location.get("longitude") if location else None
                
                utm_e_calc, utm_n_calc, _ = latlon_to_utm(lat_gps, lon_gps)

                with st.form("form_furo", clear_on_submit=True):
                    c1, c2, c3 = st.columns(3)
                    furo_id = c1.text_input("Identificação do Furo (ex: FSD-01)")
                    sonda_furo = c2.selectbox("Sonda Responsável", df_sondas["codigo"].tolist())
                    sonda_furo_id = df_sondas[df_sondas["codigo"] == sonda_furo]["id"].values[0]
                    situacao = c3.selectbox("Situação Inicial", ["Planejado", "Em Andamento", "Concluído", "Cancelado"])

                    c4, c5, c6 = st.columns(3)
                    coord_e = c4.number_input("Coordenada Este (UTM)", value=float(utm_e_calc), format="%.2f")
                    coord_n = c5.number_input("Coordenada Norte (UTM)", value=float(utm_n_calc), format="%.2f")
                    cota = c6.number_input("Cota (m)", value=0.0, step=0.5)

                    prof_plan = st.number_input("Profundidade Planejada (m)", min_value=1.0, step=1.0)

                    btn_salvar_furo = st.form_submit_button("Cadastrar Furo", type="primary", use_container_width=True)

                    if btn_salvar_furo and furo_id:
                        conn = get_connection()
                        cursor = conn.cursor()
                        try:
                            cursor.execute(
                                """
                                INSERT INTO furos (id, sonda_id, coord_e, coord_n, cota, prof_planejada, situacao)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                                """,
                                (furo_id, int(sonda_furo_id), coord_e, coord_n, cota, prof_plan, situacao)
                            )
                            conn.commit()
                            st.success(f"Furo '{furo_id}' cadastrado com sucesso!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Já existe um furo cadastrado com este ID.")
                        finally:
                            conn.close()

# ------------------------------------------------------------------------------
# 5. BOLETIM GEOLÓGICO
# ------------------------------------------------------------------------------
elif opcao == "⛏️ Boletim Geológico":
    st.title("⛏️ Boletim Geológico / Descrição de Testemunhos")
    st.caption("Registro de manobras, recuperação, RQD, descrição litológica e envio de fotos.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_furos = pd.read_sql_query("SELECT id FROM furos", conn)
        df_boletim = pd.read_sql_query("SELECT * FROM boletim_geologico ORDER BY furo_id, de_m ASC", conn)
    else:
        df_furos = pd.read_sql_query("SELECT id FROM furos WHERE sonda_id = ?", conn, params=(sonda_id_atual,))
        df_boletim = pd.read_sql_query(
            "SELECT bg.* FROM boletim_geologico bg JOIN furos f ON bg.furo_id = f.id WHERE f.sonda_id = ? ORDER BY bg.furo_id, bg.de_m ASC",
            conn,
            params=(sonda_id_atual,),
        )
    conn.close()

    tab_geo_lista, tab_geo_novo = st.tabs(["📋 Registros Geológicos", "➕ Novo Registro Geológico"])

    with tab_geo_lista:
        if not df_boletim.empty:
            st.dataframe(df_boletim, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum registro geológico encontrado.")

    with tab_geo_novo:
        if not df_furos.empty:
            with st.container(border=True):
                st.subheader("Novo Registro de Testemunho de Sondagem")
                
                furo_geo_id = st.selectbox("Selecione o Furo", df_furos["id"].tolist())
                
                c1, c2 = st.columns(2)
                de_m = c1.number_input("De (m)", min_value=0.0, step=0.1)
                ate_m = c2.number_input("Até (m)", min_value=0.0, step=0.1)

                c3, c4 = st.columns(2)
                rec_m = c3.number_input("Recuperação (m)", min_value=0.0, step=0.01)
                rqd_m = c4.number_input("RQD (m)", min_value=0.0, step=0.01)

                c5, c6 = st.columns(2)
                litologia = c5.text_input("Litologia / Tipo de Solo/Rocha")
                n_amostra = c6.text_input("Nº da Amostra")

                desc_geo = st.text_area("Descrição Geológica Detalhada")
                obs = st.text_input("Observações")
                
                foto_upload = st.file_uploader("Foto da Caixa de Testemunho", type=["jpg", "jpeg", "png"])

                if st.button("Salvar Registro Geológico", type="primary", use_container_width=True):
                    if ate_m <= de_m:
                        st.error("O valor 'Até (m)' deve ser maior que 'De (m)'.")
                    else:
                        foto_url = None
                        if foto_upload:
                            foto_url = upload_foto_supabase(foto_upload, foto_upload.name)

                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            """
                            INSERT INTO boletim_geologico 
                            (furo_id, de_m, ate_m, recuperacao_m, rqd_m, litologia, descricao_geologica, n_amostra, observacoes, foto_url)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (furo_geo_id, de_m, ate_m, rec_m, rqd_m, litologia, desc_geo, n_amostra, obs, foto_url)
                        )
                        conn.commit()
                        conn.close()

                        # Grava em paralelo no Supabase PostgreSQL se configurado
                        salvar_boletim_supabase({
                            "furo_id": furo_geo_id,
                            "de_m": de_m,
                            "ate_m": ate_m,
                            "recuperacao_m": rec_m,
                            "rqd_m": rqd_m,
                            "litologia": litologia,
                            "descricao_geologica": desc_geo,
                            "n_amostra": n_amostra,
                            "observacoes": obs,
                            "foto_url": foto_url
                        })

                        st.success("Registro geológico inserido com sucesso!")
                        st.rerun()
        else:
            st.warning("Cadastre furos antes de adicionar informações no boletim geológico.")

# ------------------------------------------------------------------------------
# 6. GESTÃO DE USUÁRIOS (APENAS ADMIN)
# ------------------------------------------------------------------------------
elif opcao == "👥 Gestão de Usuários" and perfil_atual == "Admin":
    st.title("👥 Controle de Acesso e Perfil de Usuários")
    st.caption("Administração de privilégios e associação de usuários a sondas específicas.")
    st.markdown("---")

    conn = get_connection()
    df_usuarios = pd.read_sql_query(
        "SELECT u.id, u.usuario, u.perfil, s.codigo as sonda_associada FROM usuarios u LEFT JOIN sondas s ON u.sonda_id = s.id",
        conn
    )
    df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
    conn.close()

    tab_usr_lista, tab_usr_novo = st.tabs(["📋 Usuários Cadastrados", "➕ Criar Novo Usuário"])

    with tab_usr_lista:
        st.dataframe(df_usuarios, use_container_width=True, hide_index=True)

    with tab_usr_novo:
        with st.container(border=True):
            with st.form("form_novo_usuario", clear_on_submit=True):
                st.subheader("Novo Usuário do Sistema")
                c1, c2 = st.columns(2)
                novo_user = c1.text_input("Nome de Usuário")
                nova_senha = c2.text_input("Senha", type="password")

                c3, c4 = st.columns(2)
                novo_perfil = c3.selectbox("Perfil de Acesso", ["Admin", "Geólogo", "Operador"])
                
                lista_sondas = ["Todas (Acesso Geral)"] + df_sondas["codigo"].tolist()
                sonda_vinc = c4.selectbox("Sonda Vinculada (Restrição de acesso)", lista_sondas)

                btn_criar_usr = st.form_submit_button("Cadastrar Usuário", type="primary", use_container_width=True)

                if btn_criar_usr and novo_user and nova_senha:
                    s_id = None
                    if sonda_vinc != "Todas (Acesso Geral)":
                        s_id = df_sondas[df_sondas["codigo"] == sonda_vinc]["id"].values[0]

                    conn = get_connection()
                    cursor = conn.cursor()
                    try:
                        cursor.execute(
                            "INSERT INTO usuarios (usuario, senha, perfil, sonda_id) VALUES (?, ?, ?, ?)",
                            (novo_user, hash_senha(nova_senha), novo_perfil, int(s_id) if s_id else None)
                        )
                        conn.commit()
                        st.success(f"Usuário '{novo_user}' criado com sucesso!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Nome de usuário já existe.")
                    finally:
                        conn.close()
