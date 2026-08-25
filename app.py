import base64
import hashlib
import io
import math
import sqlite3
from datetime import date

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from streamlit_geolocation import streamlit_geolocation

# IMPORTAÇÃO DOS MÓDULOS CUSTOMIZADOS
from ui_components import aplicar_estilo_customizado, render_kpi_card

# ==============================================================================
# AUTENTICAÇÃO E GESTÃO DE USUÁRIOS (RBAC)
# ==============================================================================


def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()


def get_connection():
    return sqlite3.connect("central_sondagem.db")


def criar_tabela_usuarios():
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela de usuários com vínculo de sonda_id
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

    # Garante que as tabelas do sistema existam
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
            FOREIGN KEY(furo_id) REFERENCES furos(id)
        )
    """)

    # Cria usuário Admin padrão caso o banco esteja vazio
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

    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False
        st.session_state["usuario"] = ""
        st.session_state["perfil"] = ""
        st.session_state["sonda_id"] = None

    if not st.session_state["autenticado"]:
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
        f"👤 **{st.session_state['usuario']}** ({st.session_state['perfil']})"
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
                   ROUND((recuperacao_m / (ate_m - de_m)) * 100, 1) as recuperacao_pct,
                   rqd_m, ROUND((rqd_m / (ate_m - de_m)) * 100, 1) as rqd_pct, litologia, n_amostra, descricao_geologica, observacoes
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
                   ROUND((bg.recuperacao_m / (bg.ate_m - bg.de_m)) * 100, 1) as recuperacao_pct,
                   bg.rqd_m, ROUND((bg.rqd_m / (bg.ate_m - bg.de_m)) * 100, 1) as rqd_pct, bg.litologia, bg.n_amostra, bg.descricao_geologica, bg.observacoes
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
    total_hrs_par = df_prod["horas_paradas"].sum() if not df_prod.empty else 0.0
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
# INICIALIZAÇÃO DA APLICAÇÃO
# ==============================================================================

st.set_page_config(page_title="Central de Controle - Sondagem", layout="wide")
add_bg_from_local("logo_empresa.png")
aplicar_estilo_customizado()

if not tela_login():
    st.stop()

perfil_atual = st.session_state["perfil"]
sonda_id_atual = st.session_state["sonda_id"]

# NAVEGAÇÃO STREAMLIT & SIDEBAR
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
        df_furos = pd.read_sql_query("SELECT id FROM furos", conn)
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
            "SELECT id FROM furos WHERE sonda_id = ?",
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
        if not df_sondas.empty:
            with st.container(border=True):
                with st.form("form_producao", clear_on_submit=True):
                    st.subheader("Registrar Avanço Diário")
                    c1, c2, c3 = st.columns(3)
                    data_reg = c1.date_input("Data", date.today())
                    sonda_sel = c2.selectbox("Sonda", df_sondas["codigo"].tolist())
                    lista_furos = (
                        df_furos["id"].tolist()
                        if not df_furos.empty
                        else ["Sem Furo Cadastrado"]
                    )
                    furo_sel = c3.selectbox("Furo", lista_furos)

                    c4, c5 = st.columns(2)
                    prof_in = c4.number_input(
                        "Profundidade Inicial (m)", min_value=0.0, step=0.1
                    )
                    prof_fim = c5.number_input(
                        "Profundidade Final (m)", min_value=0.0, step=0.1
                    )

                    c6, c7 = st.columns(2)
                    hrs_trab = c6.number_input(
                        "Horas Trabalhadas",
                        min_value=0.0,
                        max_value=24.0,
                        value=8.0,
                        step=0.5,
                    )
                    hrs_par = c7.number_input(
                        "Horas Paradas",
                        min_value=0.0,
                        max_value=24.0,
                        value=0.0,
                        step=0.5,
                    )
                    motivo_parada = st.text_input("Motivo da Parada")

                    btn_reg = st.form_submit_button(
                        "Lançar Produção", type="primary", use_container_width=True
                    )

                    if btn_reg and prof_fim >= prof_in:
                        s_id = int(
                            df_sondas[df_sondas["codigo"] == sonda_sel]["id"].values[0]
                        )
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            """
                            INSERT INTO producao_diaria (data, sonda_id, furo_id, prof_inicial, prof_final, horas_trabalhadas, horas_paradas, motivo_parada)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                str(data_reg),
                                s_id,
                                furo_sel,
                                prof_in,
                                prof_fim,
                                hrs_trab,
                                hrs_par,
                                motivo_parada,
                            ),
                        )
                        conn.commit()
                        conn.close()
                        st.success("Apontamento registrado!")
                        st.rerun()
        else:
            st.warning("Nenhuma sonda vinculada disponível para lançamento.")

    with tab_excluir:
        if not df_prod_full.empty:
            with st.container(border=True):
                st.subheader("Remover Apontamento")
                opcoes_prod_excluir = df_prod_full.apply(
                    lambda r: (
                        f"ID {r['id']} | Data: {r['data']} | Sonda: {r['sonda_codigo']} | Furo: {r['furo_id']} ({r['prof_inicial']}m - {r['prof_final']}m)"
                    ),
                    axis=1,
                ).tolist()
                apontamento_sel = st.selectbox(
                    "Selecione o registro para remoção:", opcoes_prod_excluir
                )

                if st.button(
                    "❌ Confirmar Exclusão", type="secondary", use_container_width=True
                ):
                    id_prod_excluir = int(apontamento_sel.split(" ")[1])
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute(
                        "DELETE FROM producao_diaria WHERE id = ?", (id_prod_excluir,)
                    )
                    conn.commit()
                    conn.close()
                    st.success("Apontamento removido com sucesso!")
                    st.rerun()

# ------------------------------------------------------------------------------
# 4. CONTROLE DE FUROS
# ------------------------------------------------------------------------------
elif opcao == "📍 Controle de Furos":
    st.title("📍 Controle de Furos de Sondagem")
    st.caption("Planejamento e dados geográficos de cada perfuração.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_furos_full = pd.read_sql_query("SELECT * FROM furos", conn)
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
    else:
        df_furos_full = pd.read_sql_query(
            "SELECT * FROM furos WHERE sonda_id = ?",
            conn,
            params=(sonda_id_atual,),
        )
        df_sondas = pd.read_sql_query(
            "SELECT id, codigo FROM sondas WHERE id = ?",
            conn,
            params=(sonda_id_atual,),
        )
    conn.close()

    tab_lista, tab_novo, tab_excluir = st.tabs(
        ["📋 Furos Cadastrados", "➕ Novo Furo", "🗑️ Excluir Furo"]
    )

    with tab_lista:
        if not df_furos_full.empty:
            st.dataframe(df_furos_full, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum furo cadastrado.")

    with tab_novo:
        if not df_sondas.empty:
            with st.container(border=True):
                st.subheader("Cadastrar Novo Furo")
                st.caption(
                    "Clique abaixo para obter as coordenadas via GPS do dispositivo."
                )

                location = streamlit_geolocation()

                if location and location.get("latitude"):
                    lat_cap = location["latitude"]
                    lon_cap = location["longitude"]
                    easting_auto, northing_auto, fuso_auto = latlon_to_utm(
                        lat_cap, lon_cap
                    )
                    st.success(
                        f"📍 GPS Capturado: Lat {lat_cap:.6f}, Lon {lon_cap:.6f} | UTM: E {easting_auto:.2f}, N {northing_auto:.2f} (Fuso {fuso_auto})"
                    )
                else:
                    easting_auto, northing_auto = 0.0, 0.0

                with st.form("form_furo", clear_on_submit=True):
                    c1, c2, c3 = st.columns(3)
                    id_furo = c1.text_input("ID do Furo (ex: F-01)")
                    sonda_furo = c2.selectbox(
                        "Sonda Alocada", df_sondas["codigo"].tolist()
                    )
                    prof_plan = c3.number_input(
                        "Profundidade Planejada (m)", min_value=1.0, step=5.0
                    )

                    c4, c5, c6 = st.columns(3)
                    coord_e = c4.number_input(
                        "Coordenada East (E)",
                        value=float(easting_auto),
                        format="%.2f",
                    )
                    coord_n = c5.number_input(
                        "Coordenada North (N)",
                        value=float(northing_auto),
                        format="%.2f",
                    )
                    cota = c6.number_input("Cota (Z)", value=0.0)

                    btn_furo = st.form_submit_button(
                        "Cadastrar Furo", type="primary", use_container_width=True
                    )

                    if btn_furo and id_furo and sonda_furo:
                        s_id = int(
                            df_sondas[df_sondas["codigo"] == sonda_furo][
                                "id"
                            ].values[0]
                        )
                        conn = get_connection()
                        cursor = conn.cursor()
                        try:
                            cursor.execute(
                                "INSERT INTO furos (id, sonda_id, coord_e, coord_n, cota, prof_planejada) VALUES (?, ?, ?, ?, ?, ?)",
                                (
                                    id_furo,
                                    s_id,
                                    coord_e,
                                    coord_n,
                                    cota,
                                    prof_plan,
                                ),
                            )
                            conn.commit()
                            st.success(f"Furo {id_furo} cadastrado!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("ID de furo já existe.")
                        finally:
                            conn.close()
        else:
            st.warning("Nenhuma sonda disponível para alocar furos.")

    with tab_excluir:
        if not df_furos_full.empty:
            with st.container(border=True):
                st.subheader("Excluir Furo do Sistema")
                furo_para_excluir = st.selectbox(
                    "Selecione o Furo:", df_furos_full["id"].tolist()
                )
                if st.button(
                    "❌ Confirmar Exclusão do Furo",
                    type="secondary",
                    use_container_width=True,
                ):
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute(
                        "DELETE FROM boletim_geologico WHERE furo_id = ?",
                        (furo_para_excluir,),
                    )
                    cursor.execute(
                        "DELETE FROM producao_diaria WHERE furo_id = ?",
                        (furo_para_excluir,),
                    )
                    cursor.execute(
                        "DELETE FROM furos WHERE id = ?", (furo_para_excluir,)
                    )
                    conn.commit()
                    conn.close()
                    st.success(f"Furo {furo_para_excluir} removido!")
                    st.rerun()

# ------------------------------------------------------------------------------
# 5. BOLETIM GEOLÓGICO
# ------------------------------------------------------------------------------
elif opcao == "⛏️ Boletim Geológico":
    st.title("⛏️ Boletim Geológico de Sondagem")
    st.caption("Descrição litológica, recuperação de testemunho e cálculo RQD.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_furos = pd.read_sql_query("SELECT id FROM furos", conn)
    else:
        df_furos = pd.read_sql_query(
            "SELECT id FROM furos WHERE sonda_id = ?",
            conn,
            params=(sonda_id_atual,),
        )

    if df_furos.empty:
        st.warning(
            "Nenhum furo associado disponível para registrar boletim geológico."
        )
        conn.close()
    else:
        furo_selecionado = st.selectbox(
            "Selecione o Furo em Operação", df_furos["id"].tolist()
        )

        df_geo = pd.read_sql_query(
            """
            SELECT id, de_m, ate_m, (ate_m - de_m) as avanco_m, recuperacao_m, 
                   ROUND((recuperacao_m / (ate_m - de_m)) * 100, 1) as recuperacao_pct,
                   rqd_m, ROUND((rqd_m / (ate_m - de_m)) * 100, 1) as rqd_pct,
                   litologia, n_amostra, descricao_geologica, observacoes
            FROM boletim_geologico WHERE furo_id = ? ORDER BY de_m ASC
        """,
            conn,
            params=(furo_selecionado,),
        )
        conn.close()

        tab_boletim, tab_novo_trecho = st.tabs(
            ["📋 Registros do Furo", "➕ Novo Trecho Geológico"]
        )

        with tab_boletim:
            if not df_geo.empty:
                st.dataframe(df_geo, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhum trecho geológico registrado para este furo.")

        with tab_novo_trecho:
            with st.container(border=True):
                with st.form("form_boletim", clear_on_submit=True):
                    st.subheader(f"Novo Trecho para {furo_selecionado}")
                    c1, c2, c3, c4 = st.columns(4)
                    de_m = c1.number_input("De (m)", min_value=0.0, step=0.5)
                    ate_m = c2.number_input("Até (m)", min_value=0.0, step=0.5)
                    rec_m = c3.number_input(
                        "Recuperação (m)", min_value=0.0, step=0.1
                    )
                    rqd_m = c4.number_input("RQD (m)", min_value=0.0, step=0.1)

                    c5, c6 = st.columns(2)
                    litologia = c5.text_input("Litologia / Rocha")
                    amostra = c6.text_input("Nº da Amostra")

                    desc = st.text_area("Descrição Geológica")
                    obs = st.text_input("Observações")

                    btn_salvar_geo = st.form_submit_button(
                        "Salvar Trecho Geológico",
                        type="primary",
                        use_container_width=True,
                    )

                    if btn_salvar_geo and ate_m > de_m:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            """
                            INSERT INTO boletim_geologico (furo_id, de_m, ate_m, recuperacao_m, rqd_m, litologia, descricao_geologica, n_amostra, observacoes)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                furo_selecionado,
                                de_m,
                                ate_m,
                                rec_m,
                                rqd_m,
                                litologia,
                                desc,
                                amostra,
                                obs,
                            ),
                        )
                        conn.commit()
                        conn.close()
                        st.success("Trecho geológico adicionado!")
                        st.rerun()

# ------------------------------------------------------------------------------
# 6. GESTÃO DE USUÁRIOS (APENAS PERFIL ADMIN)
# ------------------------------------------------------------------------------
elif opcao == "👥 Gestão de Usuários" and perfil_atual == "Admin":
    st.title("👥 Gestão de Usuários e Permissões")
    st.caption("Cadastre novos operadores e vincule cada um à sua respectiva sonda.")
    st.markdown("---")

    conn = get_connection()
    df_users = pd.read_sql_query(
        """
        SELECT u.id, u.usuario, u.perfil, s.codigo as sonda_vinculada
        FROM usuarios u LEFT JOIN sondas s ON u.sonda_id = s.id
    """,
        conn,
    )
    df_sondas_cad = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
    conn.close()

    tab_usr_lista, tab_usr_novo = st.tabs(
        ["📋 Usuários Cadastrados", "➕ Novo Usuário"]
    )

    with tab_usr_lista:
        st.dataframe(df_users, use_container_width=True, hide_index=True)

    with tab_usr_novo:
        with st.container(border=True):
            with st.form("form_novo_usuario", clear_on_submit=True):
                st.subheader("Cadastrar Novo Usuário")
                c1, c2 = st.columns(2)
                novo_user = c1.text_input("Nome de Usuário (Login)")
                nova_senha = c2.text_input("Senha", type="password")

                c3, c4 = st.columns(2)
                perfil_sel = c3.selectbox(
                    "Perfil de Acesso", ["Operador", "Geólogo", "Admin"]
                )

                sondas_opcoes = ["Nenhuma / Acesso Total (Admin)"] + (
                    df_sondas_cad["codigo"].tolist()
                    if not df_sondas_cad.empty
                    else []
                )
                sonda_vinculo = c4.selectbox(
                    "Vincular à Sonda (Campo)", sondas_opcoes
                )

                btn_cad_user = st.form_submit_button(
                    "Salvar Usuário", type="primary", use_container_width=True
                )

                if btn_cad_user and novo_user and nova_senha:
                    s_id_val = None
                    if (
                        sonda_vinculo != "Nenhuma / Acesso Total (Admin)"
                        and not df_sondas_cad.empty
                    ):
                        s_id_val = int(
                            df_sondas_cad[df_sondas_cad["codigo"] == sonda_vinculo][
                                "id"
                            ].values[0]
                        )

                    conn = get_connection()
                    cursor = conn.cursor()
                    try:
                        cursor.execute(
                            "INSERT INTO usuarios (usuario, senha, perfil, sonda_id) VALUES (?, ?, ?, ?)",
                            (
                                novo_user,
                                hash_senha(nova_senha),
                                perfil_sel,
                                s_id_val,
                            ),
                        )
                        conn.commit()
                        st.success(f"Usuário '{novo_user}' criado com sucesso!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Este nome de usuário já está em uso.")
                    finally:
                        conn.close()
