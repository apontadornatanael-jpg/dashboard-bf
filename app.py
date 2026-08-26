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
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
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
    layout="wide",
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

SUPABASE_URL = st.secrets.get(
    "SUPABASE_URL", "https://seu-projeto.supabase.co"
)
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

        return supabase.storage.from_(bucket_name).get_public_url(
            caminho_storage
        )
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
# HELPER DE EXPORTAÇÃO EXCEL AUTOMATIZADO COM DASHBOARD E GRÁFICOS
# ==============================================================================


def gerar_dashboard_excel_completo(perfil_usuario, user_sonda_id):
    output = io.BytesIO()
    wb = openpyxl.Workbook()

    # Estilos Visuais Profissionais
    COLOR_NAVY = "1B365D"
    COLOR_HEADER_BG = "2C3E50"
    COLOR_HEADER_TXT = "FFFFFF"
    COLOR_ZEBRA = "F8F9FA"
    COLOR_CARD_BG = "F4F6F7"
    COLOR_BORDER = "D9D9D9"

    fonte_titulo = Font(name="Arial", size=16, bold=True, color="FFFFFF")
    fonte_cabecalho = Font(name="Arial", size=10, bold=True, color=COLOR_HEADER_TXT)
    fonte_dados = Font(name="Arial", size=10)
    fonte_bold = Font(name="Arial", size=10, bold=True)
    fonte_kpi_num = Font(name="Arial", size=18, bold=True, color=COLOR_NAVY)
    fonte_kpi_lbl = Font(name="Arial", size=9, bold=True, color="7F8C8D")

    fill_navy = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
    fill_header = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    fill_zebra = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid")
    fill_card = PatternFill(start_color=COLOR_CARD_BG, end_color=COLOR_CARD_BG, fill_type="solid")
    fill_logo = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")

    borda_fina = Side(border_style="thin", color=COLOR_BORDER)
    borda_caixa = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)

    # Buscar dados do SQLite
    conn = get_connection()
    if perfil_usuario == "Admin":
        df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
        df_prod = pd.read_sql_query(
            "SELECT p.*, s.codigo as sonda_codigo FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id",
            conn,
        )
        df_geo = pd.read_sql_query(
            """
            SELECT bg.id, bg.furo_id, bg.de_m, bg.ate_m, (bg.ate_m - bg.de_m) as avanco_m, bg.recuperacao_m, 
                   ROUND((bg.recuperacao_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as recuperacao_pct,
                   bg.rqd_m, ROUND((bg.rqd_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as rqd_pct, bg.litologia, bg.n_amostra, bg.descricao_geologica, bg.observacoes
            FROM boletim_geologico bg ORDER BY bg.furo_id, bg.de_m ASC
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
                   bg.rqd_m, ROUND((bg.rqd_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as rqd_pct, bg.litologia, bg.n_amostra, bg.descricao_geologica, bg.observacoes
            FROM boletim_geologico bg JOIN furos f ON bg.furo_id = f.id WHERE f.sonda_id = ? ORDER BY bg.furo_id, bg.de_m ASC
        """,
            conn,
            params=(user_sonda_id,),
        )
    conn.close()

    # 1. ABA DE DASHBOARD
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_dash.views.sheetView[0].showGridLines = True

    # Espaço para Logomarca
    ws_dash.merge_cells("A1:C3")
    logo_cell = ws_dash["A1"]
    logo_cell.value = "Espaço para Logo\n[ Insira a Imagem Aqui ]"
    logo_cell.font = Font(name="Arial", size=9, italic=True, color="7F8C8D")
    logo_cell.fill = fill_logo
    logo_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_dash.merge_cells("D1:L3")
    dash_title = ws_dash["D1"]
    dash_title.value = "CENTRAL DE CONTROLE — OPERAÇÕES DE SONDAGEM"
    dash_title.font = fonte_titulo
    dash_title.fill = fill_navy
    dash_title.alignment = Alignment(horizontal="center", vertical="center")

    # Função Helper KPI Cards
    def criar_kpi_card(ws, start_col, start_row, label, formula_val, num_format=None):
        col_s = get_column_letter(start_col)
        col_e = get_column_letter(start_col + 2)
        ws.merge_cells(f"{col_s}{start_row}:{col_e}{start_row}")
        lbl_c = ws[f"{col_s}{start_row}"]
        lbl_c.value = label.upper()
        lbl_c.font = fonte_kpi_lbl
        lbl_c.fill = fill_card
        lbl_c.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(f"{col_s}{start_row+1}:{col_e}{start_row+2}")
        val_c = ws[f"{col_s}{start_row+1}"]
        val_c.value = formula_val
        val_c.font = fonte_kpi_num
        val_c.fill = fill_card
        val_c.alignment = Alignment(horizontal="center", vertical="center")
        if num_format:
            val_c.number_format = num_format

        for r in range(start_row, start_row + 3):
            for c in range(start_col, start_col + 3):
                ws.cell(row=r, column=c).border = borda_caixa

    criar_kpi_card(ws_dash, 1, 5, "Total de Registros", "=COUNTA('Registros de Sondagem'!A5:A200)")
    criar_kpi_card(ws_dash, 4, 5, "Avanço Acumulado", "=SUM('Registros de Sondagem'!H5:H200)", "#,##0.00\" m\"")
    criar_kpi_card(ws_dash, 7, 5, "Recuperação Média", "=AVERAGE('Registros de Sondagem'!I5:I200)", "0.0%")
    criar_kpi_card(ws_dash, 10, 5, "RQD Médio", "=AVERAGE('Registros de Sondagem'!J5:J200)", "0.0%")

    # Tabela Auxiliar de Agregação no Dashboard para Gráficos
    ws_dash.cell(row=10, column=1, value="Sonda / Furo").font = fonte_bold
    ws_dash.cell(row=10, column=2, value="Avanço (m)").font = fonte_bold

    litos_unicas = df_geo["litologia"].unique().tolist() if not df_geo.empty else ["Solo", "Xisto", "Gnaisse"]
    sondas_unicas = df_sondas["codigo"].unique().tolist() if not df_sondas.empty else ["Sonda-01"]

    for idx, s_cod in enumerate(sondas_unicas[:5], start=11):
        ws_dash.cell(row=idx, column=1, value=s_cod).font = fonte_dados
        c = ws_dash.cell(row=idx, column=2, value=f"=SUMIF('Registros de Sondagem'!$C$5:$C$200, A{idx}, 'Registros de Sondagem'!$H$5:$H$200)")
        c.font = fonte_dados
        c.number_format = "#,##0.00"

    chart_bar = BarChart()
    chart_bar.type = "col"
    chart_bar.style = 10
    chart_bar.title = "Avanço por Sonda (m)"
    chart_bar.y_axis.title = "Metros"
    chart_bar.legend = None
    chart_bar.width = 14
    chart_bar.height = 9

    data_bar = Reference(ws_dash, min_col=2, min_row=10, max_row=10 + len(sondas_unicas[:5]))
    cats_bar = Reference(ws_dash, min_col=1, min_row=11, max_row=10 + len(sondas_unicas[:5]))
    chart_bar.add_data(data_bar, titles_from_data=True)
    chart_bar.set_categories(cats_bar)
    ws_dash.add_chart(chart_bar, "A16")

    # Tabela 2: Distribuição por Litologia
    ws_dash.cell(row=10, column=5, value="Litologia").font = fonte_bold
    ws_dash.cell(row=10, column=6, value="Metragem (m)").font = fonte_bold

    for idx, lit in enumerate(litos_unicas[:6], start=11):
        ws_dash.cell(row=idx, column=5, value=lit).font = fonte_dados
        c = ws_dash.cell(row=idx, column=6, value=f"=SUMIF('Registros de Sondagem'!$E$5:$E$200, E{idx}, 'Registros de Sondagem'!$H$5:$H$200)")
        c.font = fonte_dados
        c.number_format = "#,##0.00"

    chart_pie = PieChart()
    chart_pie.title = "Distribuição Litológica"
    chart_pie.width = 14
    chart_pie.height = 9

    data_pie = Reference(ws_dash, min_col=6, min_row=10, max_row=10 + len(litos_unicas[:6]))
    cats_pie = Reference(ws_dash, min_col=5, min_row=11, max_row=10 + len(litos_unicas[:6]))
    chart_pie.add_data(data_pie, titles_from_data=True)
    chart_pie.set_categories(cats_pie)
    ws_dash.add_chart(chart_pie, "G16")

    # 2. ABA DE REGISTROS DE SONDAGEM
    ws_data = wb.create_sheet(title="Registros de Sondagem")
    ws_data.views.sheetView[0].showGridLines = True

    ws_data.merge_cells("A1:K2")
    title_data = ws_data["A1"]
    title_data.value = "BASE DE DADOS DE BOLETIM GEOLÓGICO DE CAMPO"
    title_data.font = fonte_titulo
    title_data.fill = fill_navy
    title_data.alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        "ID Registro", "Furo ID", "Sonda", "De (m)", "Até (m)", 
        "Recuperação (m)", "RQD (m)", "Avanço (m)", "Recuperação (%)", "RQD (%)", "Litologia"
    ]

    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws_data.cell(row=4, column=c_idx, value=h_text)
        cell.font = fonte_cabecalho
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_caixa

    row_count = 5
    if not df_geo.empty:
        for r_idx, r in df_geo.iterrows():
            ws_data.cell(row=row_count, column=1, value=r["id"]).alignment = Alignment(horizontal="center")
            ws_data.cell(row=row_count, column=2, value=str(r["furo_id"])).alignment = Alignment(horizontal="center")
            ws_data.cell(row=row_count, column=3, value=sondas_unicas[0] if sondas_unicas else "SD-01").alignment = Alignment(horizontal="center")
            ws_data.cell(row=row_count, column=4, value=r["de_m"]).number_format = "#,##0.00"
            ws_data.cell(row=row_count, column=5, value=r["ate_m"]).number_format = "#,##0.00"
            ws_data.cell(row=row_count, column=6, value=r["recuperacao_m"]).number_format = "#,##0.00"
            ws_data.cell(row=row_count, column=7, value=r["rqd_m"]).number_format = "#,##0.00"

            # Avanço Calculado Dinamicamente
            c_av = ws_data.cell(row=row_count, column=8, value=f"=E{row_count}-D{row_count}")
            c_av.number_format = "#,##0.00"

            # Percentuais Calculados por Fórmula
            c_rec_pct = ws_data.cell(row=row_count, column=9, value=f"=IFERROR(F{row_count}/H{row_count}, 0)")
            c_rec_pct.number_format = "0.0%"

            c_rqd_pct = ws_data.cell(row=row_count, column=10, value=f"=IFERROR(G{row_count}/H{row_count}, 0)")
            c_rqd_pct.number_format = "0.0%"

            ws_data.cell(row=row_count, column=11, value=r["litologia"])

            for col in range(1, 12):
                cell_item = ws_data.cell(row=row_count, column=col)
                cell_item.font = fonte_dados
                cell_item.border = borda_caixa
                if row_count % 2 == 0:
                    cell_item.fill = fill_zebra
            row_count += 1
    else:
        row_count = 6

    # Linha Totais
    ws_data.cell(row=row_count, column=1, value="TOTAL / MÉDIA").font = fonte_bold
    ws_data.cell(row=row_count, column=8, value=f"=SUM(H5:H{row_count-1})").number_format = "#,##0.00"
    ws_data.cell(row=row_count, column=9, value=f"=AVERAGE(I5:I{row_count-1})").number_format = "0.0%"
    ws_data.cell(row=row_count, column=10, value=f"=AVERAGE(J5:J{row_count-1})").number_format = "0.0%"

    for col in range(1, 12):
        cell_t = ws_data.cell(row=row_count, column=col)
        cell_t.font = fonte_bold
        cell_t.border = Border(top=Side(style="thin"), bottom=Side(style="double"))
        cell_t.fill = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")

    # 3. ABA DE TABELAS AUXILIARES
    ws_aux = wb.create_sheet(title="Tabelas Auxiliares")
    ws_aux.views.sheetView[0].showGridLines = True
    ws_aux.merge_cells("A1:B1")
    ws_aux["A1"] = "Domínios & Validações"
    ws_aux["A1"].font = Font(name="Arial", size=12, bold=True, color=COLOR_NAVY)

    ws_aux.cell(row=2, column=1, value="Litologias").font = fonte_bold
    ws_aux.cell(row=2, column=2, value="Status").font = fonte_bold

    aux_litologia = ["Solo de Alteração", "Xisto", "Quartzito", "Gnaisse", "Filito", "Itabirito"]
    aux_status = ["Operando", "Parada", "Manutenção", "Concluído"]

    max_r = max(len(aux_litologia), len(aux_status))
    for i in range(max_r):
        r_num = i + 3
        if i < len(aux_litologia):
            ws_aux.cell(row=r_num, column=1, value=aux_litologia[i]).font = fonte_dados
        if i < len(aux_status):
            ws_aux.cell(row=r_num, column=2, value=aux_status[i]).font = fonte_dados

    # Validação de Dados na aba principal
    dv_lito = DataValidation(type="list", formula1="'Tabelas Auxiliares'!$A$3:$A$8", allow_blank=True)
    ws_data.add_data_validation(dv_lito)
    dv_lito.add(f"K5:K{row_count+20}")

    # Ajustar largura de colunas automaticamente
    for ws in [ws_dash, ws_data, ws_aux]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.coordinate in ws.merged_cells:
                    continue
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    ws_dash.column_dimensions["A"].width = 16
    ws_dash.column_dimensions["B"].width = 16
    ws_dash.column_dimensions["C"].width = 16
    ws_dash.column_dimensions["D"].width = 16
    ws_dash.column_dimensions["E"].width = 18
    ws_dash.column_dimensions["F"].width = 16

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
# 1. DASHBOARD GERAL (ILUMINADO E DESTAQUE DE SONDAS)
# ------------------------------------------------------------------------------
if opcao == "Dashboard Geral":
    st.title("⚡ Painel Geral de Operações & Sondas")
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
    sondas_par = len(df_sondas[df_sondas["status"] == "Parada"])
    sondas_manut = len(df_sondas[df_sondas["status"] == "Manutenção"])

    if not df_prod.empty:
        df_hoje = df_prod[df_prod["data"] == str(date.today())]
        metros_hoje = df_hoje["prof_final"].sum() - df_hoje["prof_inicial"].sum()
        metros_acumulados = (df_prod["prof_final"] - df_prod["prof_inicial"]).sum()
        hrs_trab = df_prod["horas_trabalhadas"].sum()
        hrs_par = df_prod["horas_paradas"].sum()
        eficiencia = ((hrs_trab / (hrs_trab + hrs_par) * 100) if (hrs_trab + hrs_par) > 0 else 0)
    else:
        metros_hoje, metros_acumulados, eficiencia = 0.0, 0.0, 0.0

    st.markdown(
        """
        <style>
        .card-destaque {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            border-left: 5px solid #00f2fe;
            border-radius: 10px;
            padding: 18px;
            margin-bottom: 15px;
            box-shadow: 0 4px 15px rgba(0, 242, 254, 0.15);
            color: #ffffff;
        }
        .card-sonda-op {
            background: rgba(16, 185, 129, 0.1);
            border: 1px solid #10b981;
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 10px;
        }
        .card-sonda-par {
            background: rgba(239, 68, 68, 0.1);
            border: 1px solid #ef4444;
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 10px;
        }
        .card-sonda-manut {
            background: rgba(245, 158, 11, 0.1);
            border: 1px solid #f59e0b;
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 10px;
        }
        .status-badge {
            font-weight: bold;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">🟢 SONDAS ATIVAS</span>
                <h2 style="margin: 5px 0 0 0; color: #4ade80;">{sondas_op} <span style="font-size:16px; color:#cbd5e1;">/ {sondas_total}</span></h2>
            </div>
        """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">🎯 PRODUÇÃO HOJE</span>
                <h2 style="margin: 5px 0 0 0; color: #38bdf8;">{metros_hoje:.1f} m</h2>
            </div>
        """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">📐 TOTAL ACUMULADO</span>
                <h2 style="margin: 5px 0 0 0; color: #818cf8;">{metros_acumulados:.1f} m</h2>
            </div>
        """,
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">⚡ EFICIÊNCIA OPERACIONAL</span>
                <h2 style="margin: 5px 0 0 0; color: #f43f5e;">{eficiencia:.0f}%</h2>
            </div>
        """,
            unsafe_allow_html=True,
        )

    st.markdown("### 🚜 Monitoramento Individual de Sondas")

    if not df_sondas.empty:
        grid_cols = st.columns(3)
        for idx, row in df_sondas.iterrows():
            col_idx = idx % 3
            status = row["status"]

            if status == "Operando":
                css_class = "card-sonda-op"
                cor_status = "#10b981"
                icone = "🟢"
            elif status == "Parada":
                css_class = "card-sonda-par"
                cor_status = "#ef4444"
                icone = "🔴"
            else:
                css_class = "card-sonda-manut"
                cor_status = "#f59e0b"
                icone = "🟡"

            if not df_prod.empty:
                prod_sonda = df_prod[df_prod["sonda_id"] == row["id"]]
                m_sonda = (
                    (prod_sonda["prof_final"] - prod_sonda["prof_inicial"]).sum()
                    if not prod_sonda.empty
                    else 0.0
                )
            else:
                m_sonda = 0.0

            with grid_cols[col_idx]:
                st.markdown(
                    f"""
                    <div class="{css_class}">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <h3 style="margin:0; font-size:18px;">Sonda {row['codigo']}</h3>
                            <span class="status-badge" style="background:{cor_status}22; color:{cor_status}; border: 1px solid {cor_status};">
                                {icone} {status.upper()}
                            </span>
                        </div>
                        <p style="margin:5px 0; font-size:13px; color:#cbd5e1;"><b>Equipe:</b> {row['equipe']}</p>
                        <p style="margin:2px 0; font-size:13px; color:#cbd5e1;"><b>Projeto:</b> {row['projeto']}</p>
                        <hr style="margin:8px 0; border:0; border-top:1px solid rgba(255,255,255,0.1);">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span style="font-size:12px; color:#94a3b8;">Avanço Acumulado:</span>
                            <strong style="font-size:15px; color:#38bdf8;">{m_sonda:.1f} m</strong>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
    else:
        st.info("Nenhuma sonda disponível no momento.")

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

            st.markdown("---")
            st.subheader("🗑️ Excluir Sonda")

            opcoes_sonda = [
                f"{row['id']} - {row['codigo']} ({row['equipe']})"
                for _, row in df_sondas.iterrows()
            ]
            sonda_excluir_str = st.selectbox(
                "Selecione a Sonda para excluir:", opcoes_sonda
            )
            sonda_id_excluir = int(sonda_excluir_str.split(" - ")[0])

            if st.button("Excluir Sonda Selecionada", type="primary"):
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute(
                    "DELETE FROM sondas WHERE id = ?", (sonda_id_excluir,)
                )
                conn.commit()
                conn.close()
                st.success("Sonda excluída com sucesso!")
                st.rerun()
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
                        cursor.execute(
                            "INSERT INTO sondas (codigo, equipe, projeto, status) VALUES (?, ?, ?, ?)",
                            (codigo, equipe, projeto, status),
                        )
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
        df_prod = pd.read_sql_query(
            "SELECT p.id, p.data, s.codigo as sonda, p.furo_id, p.prof_inicial, p.prof_final, (p.prof_final - p.prof_inicial) as avanco, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id ORDER BY p.data DESC",
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
        df_prod = pd.read_sql_query(
            "SELECT p.id, p.data, s.codigo as sonda, p.furo_id, p.prof_inicial, p.prof_final, (p.prof_final - p.prof_inicial) as avanco, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada FROM producao_diaria p LEFT JOIN sondas s ON p.sonda_id = s.id WHERE p.sonda_id = ? ORDER BY p.data DESC",
            conn,
            params=(sonda_id_atual,),
        )
    conn.close()

    tab_hist, tab_novo = st.tabs(["Histórico", "Registrar Apontamento"])

    with tab_hist:
        if not df_prod.empty:
            st.dataframe(df_prod, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("🗑️ Excluir Registro de Apontamento")

            opcoes_prod = [
                f"ID: {row['id']} | Data: {row['data']} | Furo: {row['furo_id']} | Avanço: {row['avanco']}m"
                for _, row in df_prod.iterrows()
            ]
            prod_excluir_str = st.selectbox(
                "Selecione o Apontamento para excluir:", opcoes_prod
            )
            prod_id_excluir = int(
                prod_excluir_str.split(" | ")[0].replace("ID: ", "")
            )

            if st.button("Excluir Apontamento Selecionado", type="primary"):
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute(
                    "DELETE FROM producao_diaria WHERE id = ?",
                    (prod_id_excluir,),
                )
                conn.commit()
                conn.close()
                st.success("Apontamento excluído com sucesso!")
                st.rerun()
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
                p_ini = c4.number_input(
                    "Prof. Inicial (m)", min_value=0.0, step=0.1
                )
                p_fin = c5.number_input(
                    "Prof. Final (m)", min_value=0.0, step=0.1
                )
                h_trab = c6.number_input(
                    "Horas Trabalhadas", min_value=0.0, step=0.5
                )

                c7, c8 = st.columns(2)
                h_par = c7.number_input(
                    "Horas Paradas", min_value=0.0, step=0.5
                )
                motivo = c8.text_input("Motivo Parada")

                if st.form_submit_button(
                    "Salvar Apontamento", type="primary"
                ):
                    s_id = df_sondas[df_sondas["codigo"] == sonda_sel][
                        "id"
                    ].values[0]
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute(
                        "INSERT INTO producao_diaria (data, sonda_id, furo_id, prof_inicial, prof_final, horas_trabalhadas, horas_paradas, motivo_parada) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        (
                            dt,
                            int(s_id),
                            furo_sel,
                            p_ini,
                            p_fin,
                            h_trab,
                            h_par,
                            motivo,
                        ),
                    )
                    conn.commit()
                    conn.close()
                    st.success("Dados de produção salvos com sucesso!")
                    st.rerun()
        else:
            st.warning(
                "É necessário possuir furos e sondas vinculados para registrar o apontamento."
            )

# ------------------------------------------------------------------------------
# 4. CONTROLE DE FUROS (COM CAPTURA DE GPS AUTOMÁTICO)
# ------------------------------------------------------------------------------
elif opcao == "Controle de Furos":
    st.title("Controle de Furos de Sondagem")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
        df_furos = pd.read_sql_query(
            "SELECT f.*, s.codigo as sonda_codigo FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id",
            conn,
        )
    else:
        df_sondas = pd.read_sql_query(
            "SELECT id, codigo FROM sondas WHERE id = ?",
            conn,
            params=(sonda_id_atual,),
        )
        df_furos = pd.read_sql_query(
            "SELECT f.*, s.codigo as sonda_codigo FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id WHERE f.sonda_id = ?",
            conn,
            params=(sonda_id_atual,),
        )
    conn.close()

    tab_furos_list, tab_furos_novo = st.tabs(
        ["Furos Cadastrados", "Novo Furo"]
    )

    with tab_furos_list:
        if not df_furos.empty:
            st.dataframe(df_furos, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("🗑️ Excluir Furo")

            opcoes_furo = df_furos["id"].tolist()
            furo_excluir = st.selectbox(
                "Selecione o Furo para excluir:", opcoes_furo
            )

            if st.button("Excluir Furo Selecionado", type="primary"):
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute(
                    "DELETE FROM furos WHERE id = ?", (furo_excluir,)
                )
                conn.commit()
                conn.close()
                st.success(f"Furo '{furo_excluir}' excluído com sucesso!")
                st.rerun()
        else:
            st.info("Nenhum furo cadastrado.")

    with tab_furos_novo:
        if not df_sondas.empty:
            st.subheader("Captura de Localização via GPS")
            st.caption(
                "Clique no botão abaixo para capturar as coordenadas exatas de onde você está no campo:"
            )

            location = streamlit_geolocation()

            lat_auto = location.get("latitude") if location else None
            lon_auto = location.get("longitude") if location else None
            alt_auto = location.get("altitude") if location else None

            if lat_auto and lon_auto:
                st.success(
                    f"📍 Coordenadas capturadas: Lat {lat_auto:.6f}, Lon {lon_auto:.6f}"
                )
            else:
                st.info(
                    "Clique no ícone de GPS acima para obter as coordenadas automaticamente ou preencha manualmente."
                )

            with st.form("form_furo", clear_on_submit=True):
                c1, c2 = st.columns(2)
                furo_id = c1.text_input("Identificação do Furo (ex: F-01)")
                sonda_sel = c2.selectbox(
                    "Sonda Responsável", df_sondas["codigo"].tolist()
                )

                c3, c4, c5 = st.columns(3)
                coord_e = c3.number_input(
                    "Longitude / Easting",
                    value=float(lon_auto) if lon_auto is not None else 0.0,
                    format="%.6f",
                )
                coord_n = c4.number_input(
                    "Latitude / Northing",
                    value=float(lat_auto) if lat_auto is not None else 0.0,
                    format="%.6f",
                )
                cota = c5.number_input(
                    "Cota (m)",
                    value=float(alt_auto) if alt_auto is not None else 0.0,
                    step=0.1,
                )

                c6, c7 = st.columns(2)
                p_plan = c6.number_input(
                    "Prof. Planejada (m)", min_value=0.0, step=1.0
                )
                situacao = c7.selectbox(
                    "Situação",
                    ["Planejado", "Em Andamento", "Concluído", "Cancelado"],
                )

                if st.form_submit_button("Salvar Furo", type="primary"):
                    if furo_id:
                        s_id = df_sondas[
                            df_sondas["codigo"] == sonda_sel
                        ]["id"].values[0]
                        conn = get_connection()
                        cursor = conn.cursor()
                        try:
                            cursor.execute(
                                "INSERT INTO furos (id, sonda_id, coord_e, coord_n, cota, prof_planejada, situacao) VALUES (?, ?, ?, ?, ?, ?, ?)",
                                (
                                    furo_id,
                                    int(s_id),
                                    coord_e,
                                    coord_n,
                                    cota,
                                    p_plan,
                                    situacao,
                                ),
                            )
                            conn.commit()
                            st.success(
                                f"Furo '{furo_id}' cadastrado com sucesso!"
                            )
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error(
                                f"Furo '{furo_id}' já está cadastrado."
                            )
                        finally:
                            conn.close()
                    else:
                        st.error("Preencha a identificação do furo.")
