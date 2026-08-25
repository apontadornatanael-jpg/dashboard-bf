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
# INICIALIZAÇÃO DA APLICAÇÃO E SESSION STATE
# ==============================================================================

st.set_page_config(
    page_title="Central de Controle - Sondagem",
    page_icon=None,
    layout="wide"
)

# Estilo CSS para ocultar elementos de marca, ícones nativos e menus do Streamlit
hide_streamlit_style = """
    <style>
    /* Oculta o menu de três pontos do Streamlit */
    #MainMenu {visibility: hidden;}
    
    /* Oculta o cabeçalho padrão e a barra superior */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    
    /* Oculta o rodapé e o painel 'Manage app' */
    footer {visibility: hidden;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    [data-testid="stDecoration"] {display: none;}
    
    /* Oculta ícones SVG internos de menus/navegação do Streamlit */
    [data-testid="stSidebarNav"] svg {display: none;}
    button[title="View fullscreen"] {display: none;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

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
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def upload_foto_supabase(file_buffer, nome_arquivo):
    try:
        supabase = get_supabase_client()
        bucket_name = "boletins-fotos"
        ext = nome_arquivo.split(".")[-1]
        caminho_storage = f"boletins/{uuid.uuid4().hex}.{ext}"

        supabase.storage.from_(bucket_name).upload(
            caminho_storage,
            file_buffer.getvalue(),
            file_options={"content-type": f"image/{ext}"},
        )

        return supabase.storage.from_(bucket_name).get_public_url(caminho_storage)
    except Exception as e:
        st.error(f"Erro ao enviar imagem para o Supabase Storage: {e}")
        return None


def salvar_boletim_supabase(dados_boletim):
    try:
        supabase = get_supabase_client()
        return supabase.table("boletim_geologico").insert(dados_boletim).execute()
    except Exception as e:
        st.error(f"Erro ao salvar boletim no Supabase: {e}")
        return None


# ==============================================================================
# BANCO DE DADOS & AUTENTICAÇÃO
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
            st.subheader("Acesso ao Sistema")
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
        f"**{st.session_state.get('usuario', '')}** ({st.session_state.get('perfil', '')})"
    )
    if st.sidebar.button("Sair / Logout"):
        st.session_state["autenticado"] = False
        st.session_state["usuario"] = ""
        st.session_state["perfil"] = ""
        st.session_state["sonda_id"] = None
        st.rerun()


# ==============================================================================
# HELPER DE EXPORTAÇÃO EXCEL
# ==============================================================================


def gerar_dashboard_excel_completo(perfil_usuario, user_sonda_id):
    output = io.BytesIO()
    wb = openpyxl.Workbook()

    cor_cabecalho = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    cor_titulo = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    cor_card = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    cor_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    fonte_titulo = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fonte_card_num = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    fonte_card_lbl = Font(name="Calibri", size=9, italic=True, color="595959")
    fonte_dados = Font(name="Calibri", size=10)

    borda_fina = Side(border_style="thin", color="D9D9D9")
    borda_caixa = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)

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
        df_sondas = pd.read_sql_query("SELECT * FROM sondas WHERE id = ?", conn, params=(user_sonda_id,))
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
    ws_dash.title = "Dashboard Executivo"
    ws_dash.views.sheetView[0].showGridLines = True
    ws_dash.merge_cells("A1:H1")
    ws_dash["A1"] = "DASHBOARD EXECUTIVO — CONTROLE INTEGRADO DE SONDAGEM"
    ws_dash["A1"].font = fonte_titulo
    ws_dash["A1"].fill = cor_titulo
    ws_dash["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[1].height = 40

    total_sondas = len(df_sondas)
    sondas_op = len(df_sondas[df_sondas["status"] == "Operando"])
    total_metros = (df_prod["prof_final"] - df_prod["prof_inicial"]).sum() if not df_prod.empty else 0.0
    total_hrs_trab = df_prod["horas_trabalhadas"].sum() if not df_prod.empty else 0.0
    total_hrs_par = df_prod["horas_paradas"].sum() if not df_prod.empty else 0.0
    eficiencia = ((total_hrs_trab / (total_hrs_trab + total_hrs_par) * 100) if (total_hrs_trab + total_hrs_par) > 0 else 0.0)

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

    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# INICIALIZAÇÃO DA INTERFACE
# ==============================================================================

aplicar_estilo_customizado()

if not tela_login():
    st.stop()

perfil_atual = st.session_state.get("perfil", "")
sonda_id_atual = st.session_state.get("sonda_id", None)

# ==============================================================================
# SIDEBAR E NAVEGAÇÃO POR PERFIL
# ==============================================================================

st.sidebar.title("CENTRAL DE CONTROLE")
botao_logout()
st.sidebar.markdown("---")

# Definição das opções permitidas baseadas no tipo de usuário
if perfil_atual == "Admin":
    opcoes_menu = [
        "Dashboard Geral",
        "Cadastro de Sondas",
        "Apontamento Diário",
        "Controle de Furos",
        "Boletim Geológico",
        "Gestão de Usuários",
    ]
else:
    opcoes_menu = [
        "Dashboard Geral",
        "Apontamento Diário",
        "Controle de Furos",
        "Boletim Geológico",
    ]

opcao = st.sidebar.radio("Navegação", opcoes_menu)

excel_mestre = gerar_dashboard_excel_completo(perfil_atual, sonda_id_atual)
st.sidebar.markdown("---")
st.sidebar.download_button(
    label="Exportar Relatório",
    data=excel_mestre,
    file_name=f"Relatorio_Sondagem_{date.today()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

# ==============================================================================
# COMPONENTES DE TELA DA APLICAÇÃO
# ==============================================================================

# ------------------------------------------------------------------------------
# 1. DASHBOARD GERAL
# ------------------------------------------------------------------------------
if opcao == "Dashboard Geral":
    st.title("Painel Geral de Operações")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
        df_prod = pd.read_sql_query(
            "SELECT p.*, s.codigo as sonda_codigo FROM producao_diaria p JOIN sondas s ON p.sonda_id = s.id", 
            conn
        )
    else:
        df_sondas = pd.read_sql_query(
            "SELECT * FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,)
        )
        df_prod = pd.read_sql_query(
            "SELECT p.*, s.codigo as sonda_codigo FROM producao_diaria p JOIN sondas s ON p.sonda_id = s.id WHERE p.sonda_id = ?", 
            conn, 
            params=(sonda_id_atual,)
        )
    conn.close()

    sondas_total = len(df_sondas)
    sondas_op = len(df_sondas[df_sondas["status"] == "Operando"])

    if not df_prod.empty:
        df_hoje = df_prod[df_prod["data"] == str(date.today())]
        metros_hoje = df_hoje["prof_final"].sum() - df_hoje["prof_inicial"].sum()
        metros_acumulados = (df_prod["prof_final"] - df_prod["prof_inicial"]).sum()
        hrs_trab = df_prod["horas_trabalhadas"].sum()
        hrs_par = df_prod["horas_paradas"].sum()
        eficiencia = ((hrs_trab / (hrs_trab + hrs_par) * 100) if (hrs_trab + hrs_par) > 0 else 0)
    else:
        metros_hoje, metros_acumulados, eficiencia = 0.0, 0.0, 0.0

    # Apenas os 4 indicadores KPI principais
    c1, c2, c3, c4 = st.columns(4)
    with c1: 
        render_kpi_card("Sondas Vinculadas", f"{sondas_op}/{sondas_total}")
    with c2: 
        render_kpi_card("Produção Hoje", f"{metros_hoje:.1f} m")
    with c3: 
        render_kpi_card("Total Acumulado", f"{metros_acumulados:.1f} m")
    with c4: 
        render_kpi_card("Eficiência", f"{eficiencia:.0f}%")

# ------------------------------------------------------------------------------
# 2. GESTÃO DE SONDAS (EXCLUSIVO ADMIN)
# ------------------------------------------------------------------------------
elif opcao == "Cadastro de Sondas" and perfil_atual == "Admin":
    st.title("Gestão Central de Sondas")
    st.markdown("---")

    conn = get_connection()
    df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
    conn.close()

    tab_lista, tab_novo = st.tabs(["Sondas Cadastradas", "Nova Sonda"])

    with tab_lista:
        if not df_sondas.empty:
            st.dataframe(df_sondas, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma sonda cadastrada.")

    with tab_novo:
        with st.form("form_sonda", clear_on_submit=True):
            codigo = st.text_input("Código da Sonda (ex: SD-01)")
            equipe = st.text_input("Equipe Responsável")
            projeto = st.text_input("Projeto / Frente")
            status = st.selectbox("Status", ["Operando", "Parada", "Manutenção"])
            if st.form_submit_button("Cadastrar", type="primary"):
                if codigo and equipe and projeto:
                    conn = get_connection()
                    cursor = conn.cursor()
                    try:
                        cursor.execute("INSERT INTO sondas (codigo, equipe, projeto, status) VALUES (?, ?, ?, ?)", (codigo, equipe, projeto, status))
                        conn.commit()
                        st.success("Sonda cadastrada com sucesso!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Sonda já cadastrada.")
                    finally:
                        conn.close()

# ------------------------------------------------------------------------------
# 3. APONTAMENTO DIÁRIO
# ------------------------------------------------------------------------------
elif opcao == "Apontamento Diário":
    st.title("Apontamento Diário de Produção")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
        df_furos = pd.read_sql_query("SELECT id, sonda_id FROM furos", conn)
        df_prod = pd.read_sql_query("SELECT p.id, p.data, s.codigo as sonda, p.furo_id, p.prof_inicial, p.prof_final, (p.prof_final - p.prof_inicial) as avanco, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id ORDER BY p.data DESC", conn)
    else:
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,))
        df_furos = pd.read_sql_query("SELECT id, sonda_id FROM furos WHERE sonda_id = ?", conn, params=(sonda_id_atual,))
        df_prod = pd.read_sql_query("SELECT p.id, p.data, s.codigo as sonda, p.furo_id, p.prof_inicial, p.prof_final, (p.prof_final - p.prof_inicial) as avanco, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id WHERE p.sonda_id = ? ORDER BY p.data DESC", conn, params=(sonda_id_atual,))
    conn.close()

    tab_hist, tab_novo = st.tabs(["Histórico", "Registrar Apontamento"])

    with tab_hist:
        if not df_prod.empty:
            st.dataframe(df_prod, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum registro encontrado.")

    with tab_novo:
        if not df_sondas.empty and not df_furos.empty:
            with st.form("form_prod", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                dt = c1.date_input("Data", date.today())
                sonda_sel = c2.selectbox("Sonda", df_sondas["codigo"].tolist())
                furo_sel = c3.selectbox("Furo", df_furos["id"].tolist())

                c4, c5, c6 = st.columns(3)
                p_ini = c4.number_input("Prof. Inicial (m)", min_value=0.0, step=0.1)
                p_fin = c5.number_input("Prof. Final (m)", min_value=0.0, step=0.1)
                h_trab = c6.number_input("Horas Trabalhadas", min_value=0.0, step=0.5)

                c7, c8 = st.columns(2)
                h_par = c7.number_input("Horas Paradas", min_value=0.0, step=0.5)
                motivo = c8.text_input("Motivo Parada")

                if st.form_submit_button("Salvar Apontamento", type="primary"):
                    s_id = df_sondas[df_sondas["codigo"] == sonda_sel]["id"].values[0]
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute(
                        "INSERT INTO producao_diaria (data, sonda_id, furo_id, prof_inicial, prof_final, horas_trabalhadas, horas_paradas, motivo_parada) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        (dt, int(s_id), furo_sel, p_ini, p_fin, h_trab, h_par, motivo)
                    )
                    conn.commit()
                    conn.close()
                    st.success("Dados de produção salvos com sucesso!")
                    st.rerun()
        else:
            st.warning("É necessário possuir furos e sondas vinculados para registrar o apontamento.")

# ------------------------------------------------------------------------------
# 4. CONTROLE DE FUROS (COM CAPTURA DE GPS AUTOMÁTICO)
# ------------------------------------------------------------------------------
elif opcao == "Controle de Furos":
    st.title("Controle de Furos de Sondagem")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
        df_furos = pd.read_sql_query("SELECT f.*, s.codigo as sonda_codigo FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id", conn)
    else:
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,))
        df_furos = pd.read_sql_query("SELECT f.*, s.codigo as sonda_codigo FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id WHERE f.sonda_id = ?", conn, params=(sonda_id_atual,))
    conn.close()

    tab_furos_list, tab_furos_novo = st.tabs(["Furos Cadastrados", "Novo Furo"])

    with tab_furos_list:
        if not df_furos.empty:
            st.dataframe(df_furos, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum furo cadastrado.")

    with tab_furos_novo:
        if not df_sondas.empty:
            st.subheader("Captura de Localização via GPS")
            st.caption("Clique no botão abaixo para capturar as coordenadas exatas de onde você está no campo:")
            
            # Captura a localização atual via navegador/dispositivo
            location = streamlit_geolocation()

            lat_auto = location.get("latitude") if location else None
            lon_auto = location.get("longitude") if location else None
            alt_auto = location.get("altitude") if location else None

            if lat_auto and lon_auto:
                st.success(f"📍 Coordenadas capturadas: Lat {lat_auto:.6f}, Lon {lon_auto:.6f}")
            else:
                st.info("Clique no ícone de GPS acima para obter as coordenadas automaticamente ou preencha manualmente.")

            with st.form("form_furo", clear_on_submit=True):
                c1, c2 = st.columns(2)
                furo_id = c1.text_input("Identificação do Furo (ex: F-01)")
                sonda_sel = c2.selectbox("Sonda Responsável", df_sondas["codigo"].tolist())

                c3, c4, c5 = st.columns(3)
                coord_e = c3.number_input(
                    "Longitude / Easting", 
                    value=float(lon_auto) if lon_auto is not None else 0.0, 
                    format="%.6f"
                )
                coord_n = c4.number_input(
                    "Latitude / Northing", 
                    value=float(lat_auto) if lat_auto is not None else 0.0, 
                    format="%.6f"
                )
                cota = c5.number_input(
                    "Cota / Altitude (Z)", 
                    value=float(alt_auto) if alt_auto is not None else 0.0, 
                    format="%.2f"
                )

                c6, c7 = st.columns(2)
                prof_plan = c6.number_input("Prof. Planejada (m)", min_value=0.0, step=1.0)
                situacao = c7.selectbox("Situação", ["Planejado", "Em Andamento", "Concluído", "Cancelado"])

                if st.form_submit_button("Cadastrar Furo", type="primary"):
                    s_id = df_sondas[df_sondas["codigo"] == sonda_sel]["id"].values[0]
                    conn = get_connection()
                    cursor = conn.cursor()
                    try:
                        cursor.execute(
                            "INSERT INTO furos (id, sonda_id, coord_e, coord_n, cota, prof_planejada, situacao) VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (furo_id, int(s_id), coord_e, coord_n, cota, prof_plan, situacao)
                        )
                        conn.commit()
                        st.success("Furo adicionado com sucesso!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Identificação do Furo já existe.")
                    finally:
                        conn.close()

# ------------------------------------------------------------------------------
# 5. BOLETIM GEOLÓGICO
# ------------------------------------------------------------------------------
elif opcao == "Boletim Geológico":
    st.title("Boletim Geológico de Campo")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_furos = pd.read_sql_query("SELECT id FROM furos", conn)
        df_geo = pd.read_sql_query("SELECT * FROM boletim_geologico ORDER BY furo_id, de_m ASC", conn)
    else:
        df_furos = pd.read_sql_query("SELECT id FROM furos WHERE sonda_id = ?", conn, params=(sonda_id_atual,))
        df_geo = pd.read_sql_query("SELECT bg.* FROM boletim_geologico bg JOIN furos f ON bg.furo_id = f.id WHERE f.sonda_id = ? ORDER BY bg.furo_id, bg.de_m ASC", conn, params=(sonda_id_atual,))
    conn.close()

    tab_bg_hist, tab_bg_novo = st.tabs(["Registros Geológicos", "Registrar Intervalo"])

    with tab_bg_hist:
        if not df_geo.empty:
            st.dataframe(df_geo, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum boletim registrado.")

    with tab_bg_novo:
        if not df_furos.empty:
            lista_furos = df_furos["id"].tolist()
            
            # Seleção do furo fora do formulário para atualização em tempo real
            furo_sel = st.selectbox("Selecione o Furo para Registro", lista_furos)

            # Busca a última profundidade gravada (ate_m) para este furo
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT MAX(ate_m) FROM boletim_geologico WHERE furo_id = ?", (furo_sel,))
            ultimo_ate = cursor.fetchone()[0]
            conn.close()

            de_auto = float(ultimo_ate) if ultimo_ate is not None else 0.0

            # Entradas de profundidade fora da form para cálculo automático dinâmico do Avanço e Recuperação %
            st.subheader("Intervalo e Métricas de Avanço")
            col_de, col_ate, col_avanco = st.columns(3)
            
            de_m = col_de.number_input("De (m)", min_value=0.0, value=de_auto, step=0.1, key="de_m_input")
            ate_m = col_ate.number_input("Até (m)", min_value=de_m, value=max(de_m, de_auto), step=0.1, key="ate_m_input")
            
            # Cálculo do Avanço Automático
            avanco_m = round(ate_m - de_m, 2)
            col_avanco.metric("Avanço Automático (m)", f"{avanco_m:.2f} m")

            with st.form("form_bg", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                rec_m = c1.number_input("Recuperação (m)", min_value=0.0, max_value=float(avanco_m) if avanco_m > 0 else 999.0, step=0.01)
                
                # Cálculo da Recuperação em %
                rec_pct = (rec_m / avanco_m * 100) if avanco_m > 0 else 0.0
                c2.text_input("Recuperação (%)", value=f"{rec_pct:.1f}%", disabled=True)
                
                rqd_m = c3.number_input("RQD (m)", min_value=0.0, max_value=float(avanco_m) if avanco_m > 0 else 999.0, step=0.01)

                c4, c5 = st.columns(2)
                litologia = c4.text_input("Litologia")
                n_amostra = c5.text_input("Nº Amostra")

                foto = st.file_uploader("Foto da Caixa de Testemunho", type=["jpg", "png", "jpeg"])
                desc = st.text_area("Descrição Geológica")
                obs = st.text_area("Observações")

                if st.form_submit_button("Salvar Boletim", type="primary"):
                    if ate_m <= de_m and avanco_m == 0:
                        st.error("A profundidade final 'Até (m)' deve ser maior que a profundidade inicial 'De (m)'.")
                    else:
                        foto_url = upload_foto_supabase(foto, foto.name) if foto else ""
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            """
                            INSERT INTO boletim_geologico 
                            (furo_id, de_m, ate_m, recuperacao_m, rqd_m, litologia, descricao_geologica, n_amostra, observacoes, foto_url)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (furo_sel, de_m, ate_m, rec_m, rqd_m, litologia, desc, n_amostra, obs, foto_url)
                        )
                        conn.commit()
                        conn.close()

                        # Sincronização com o Supabase
                        salvar_boletim_supabase({
                            "furo_id": furo_sel,
                            "de_m": de_m,
                            "ate_m": ate_m,
                            "recuperacao_m": rec_m,
                            "rqd_m": rqd_m,
                            "litologia": litologia,
                            "descricao_geologica": desc,
                            "n_amostra": n_amostra,
                            "observacoes": obs,
                            "foto_url": foto_url
                        })

                        st.success("Boletim Geológico salvo com sucesso!")
                        st.rerun()
        else:
            st.warning("Nenhum furo cadastrado para registrar boletim.")
# ------------------------------------------------------------------------------
# 6. GESTÃO DE USUÁRIOS (EXCLUSIVO ADMIN)
# ------------------------------------------------------------------------------
elif opcao == "Gestão de Usuários" and perfil_atual == "Admin":
    st.title("Vinculação de Usuários e Sondas")
    st.markdown("---")

    conn = get_connection()
    df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
    df_users = pd.read_sql_query("SELECT u.id, u.usuario, u.perfil, s.codigo as sonda_vinculada FROM usuarios u LEFT JOIN sondas s ON u.sonda_id = s.id", conn)
    conn.close()

    st.subheader("Usuários Cadastrados")
    st.dataframe(df_users, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Cadastrar / Vincular Novo Usuário")
    with st.form("form_usr", clear_on_submit=True):
        c1, c2 = st.columns(2)
        u_nome = c1.text_input("Nome do Usuário")
        u_senha = c2.text_input("Senha", type="password")

        c3, c4 = st.columns(2)
        u_perfil = c3.selectbox("Perfil de Acesso", ["Operador", "Geólogo", "Admin"])
        
        opcoes_sondas = ["Sem Vinculo"] + df_sondas["codigo"].tolist()
        u_sonda = c4.selectbox("Sonda Autorizada", opcoes_sondas)

        if st.form_submit_button("Cadastrar Usuário", type="primary"):
            if u_nome and u_senha:
                s_id = None
                if u_sonda != "Sem Vinculo":
                    s_id = int(df_sondas[df_sondas["codigo"] == u_sonda]["id"].values[0])

                conn = get_connection()
                cursor = conn.cursor()
                try:
                    cursor.execute(
                        "INSERT INTO usuarios (usuario, senha, perfil, sonda_id) VALUES (?, ?, ?, ?)",
                        (u_nome, hash_senha(u_senha), u_perfil, s_id)
                    )
                    conn.commit()
                    st.success(f"Usuário {u_nome} criado com sucesso!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Nome de usuário já existente.")
                finally:
                    conn.close()
