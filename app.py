import base64
import hashlib
import io
import math
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

    /* Ajustes leves para uso em celular */
    @media (max-width: 768px) {
        .block-container {padding: 1rem 0.7rem 4rem 0.7rem;}
        [data-testid="stHorizontalBlock"] {gap: 0.5rem;}
        button {min-height: 44px;}
    }
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
if "navegacao" not in st.session_state:
    st.session_state["navegacao"] = "Dashboard Geral"


def botao_voltar_dashboard():
    """Botão simples e grande para facilitar a navegação no celular."""
    if st.button("⬅️ Voltar ao Dashboard", use_container_width=True):
        st.session_state["navegacao"] = "Dashboard Geral"
        st.rerun()


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


def get_supabase():
    return get_supabase_client()


def _rows(response):
    return response.data if response and response.data else []


def table_df(table_name, columns="*"):
    """Lê uma tabela do Supabase e devolve DataFrame."""
    try:
        data = _rows(get_supabase().table(table_name).select(columns).execute())
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Erro ao consultar a tabela '{table_name}' no Supabase: {e}")
        return pd.DataFrame()


def insert_row(table_name, data):
    return get_supabase().table(table_name).insert(data).execute()


def update_row(table_name, data, column, value):
    return get_supabase().table(table_name).update(data).eq(column, value).execute()


def delete_row(table_name, column, value):
    return get_supabase().table(table_name).delete().eq(column, value).execute()


def criar_tabela_usuarios():
    """
    O banco agora fica 100% no Supabase.
    As tabelas devem ser criadas uma única vez usando o arquivo
    supabase_schema.sql que acompanha este projeto.
    """
    try:
        supabase = get_supabase()
        resposta = supabase.table("usuarios").select("id").limit(1).execute()
        if not resposta.data:
            try:
                supabase.table("usuarios").insert({
                    "usuario": "admin",
                    "senha": hash_senha("admin123"),
                    "perfil": "Admin",
                    "sonda_id": None,
                }).execute()
            except Exception as e:
                # Se o admin já existir, não interrompe a inicialização.
                if "duplicate" not in str(e).lower() and "unique" not in str(e).lower():
                    st.warning(f"Não foi possível criar o usuário admin automaticamente: {e}")
    except Exception as e:
        st.error(
            "Não foi possível acessar o Supabase. Verifique SUPABASE_URL, "
            f"SUPABASE_KEY e se as tabelas foram criadas. Detalhe: {e}"
        )


def verificar_login(usuario, senha):
    try:
        resultado = (
            get_supabase()
            .table("usuarios")
            .select("perfil, sonda_id")
            .eq("usuario", usuario)
            .eq("senha", hash_senha(senha))
            .limit(1)
            .execute()
        )
        dados = _rows(resultado)
        return (dados[0]["perfil"], dados[0].get("sonda_id")) if dados else None
    except Exception as e:
        st.error(f"Erro ao verificar login no Supabase: {e}")
        return None


def obter_dados_dashboard(perfil_usuario, user_sonda_id):
    df_sondas = table_df("sondas")
    df_prod = table_df("producao_diaria")
    df_geo = table_df("boletim_geologico")

    if df_prod.empty:
        df_prod = pd.DataFrame(columns=[
            "id", "data", "sonda_id", "furo_id", "prof_inicial",
            "prof_final", "horas_trabalhadas", "horas_paradas", "motivo_parada"
        ])
    if df_sondas.empty:
        df_sondas = pd.DataFrame(columns=["id", "codigo", "equipe", "projeto", "status"])
    if df_geo.empty:
        df_geo = pd.DataFrame(columns=[
            "id", "furo_id", "de_m", "ate_m", "recuperacao_m", "rqd_m",
            "litologia", "descricao_geologica", "n_amostra", "observacoes", "foto_url"
        ])

    if perfil_usuario != "Admin":
        df_sondas = df_sondas[df_sondas["id"] == user_sonda_id]
        df_prod = df_prod[df_prod["sonda_id"] == user_sonda_id]

        df_furos = table_df("furos")
        if not df_furos.empty:
            df_furos = df_furos[df_furos["sonda_id"] == user_sonda_id]
            ids_furos = set(df_furos["id"].tolist())
            df_geo = df_geo[df_geo["furo_id"].isin(ids_furos)]
        else:
            df_geo = df_geo.iloc[0:0]

    if not df_prod.empty and not df_sondas.empty:
        df_prod = df_prod.merge(
            df_sondas[["id", "codigo"]],
            left_on="sonda_id", right_on="id", how="left", suffixes=("", "_sonda")
        )
        df_prod.rename(columns={"codigo": "sonda_codigo"}, inplace=True)

    if not df_geo.empty:
        df_geo = df_geo.copy()
        df_geo["avanco_m"] = df_geo["ate_m"] - df_geo["de_m"]
        df_geo["recuperacao_pct"] = (
            df_geo["recuperacao_m"] / df_geo["avanco_m"].replace(0, pd.NA) * 100
        ).round(1)
        df_geo["rqd_pct"] = (
            df_geo["rqd_m"] / df_geo["avanco_m"].replace(0, pd.NA) * 100
        ).round(1)

    return df_sondas, df_prod, df_geo

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

    df_sondas, df_prod, df_geo = obter_dados_dashboard(perfil_usuario, user_sonda_id)

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

opcao = st.sidebar.radio("Navegação", opcoes_menu, key="navegacao")

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

# ==============================================================================
# COMPONENTES DE TELA DA APLICAÇÃO
# ==============================================================================

# ------------------------------------------------------------------------------
# 1. DASHBOARD GERAL (ILUMINADO E DESTAQUE DE SONDAS)
# ------------------------------------------------------------------------------
if opcao == "Dashboard Geral":
    st.title("⚡ Painel Geral de Operações & Sondas")
    st.markdown("---")

    df_sondas, df_prod, _ = obter_dados_dashboard(perfil_atual, sonda_id_atual)

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

    # Estilo CSS para Cards Iluminados de KPIs e Sondas
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
        unsafe_allow_html=True
    )

    # Indicadores KPI com Iluminação de Destaque
    c1, c2, c3, c4 = st.columns(4)
    with c1: 
        st.markdown(f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">🟢 SONDAS ATIVAS</span>
                <h2 style="margin: 5px 0 0 0; color: #4ade80;">{sondas_op} <span style="font-size:16px; color:#cbd5e1;">/ {sondas_total}</span></h2>
            </div>
        """, unsafe_allow_html=True)
    with c2: 
        st.markdown(f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">🎯 PRODUÇÃO HOJE</span>
                <h2 style="margin: 5px 0 0 0; color: #38bdf8;">{metros_hoje:.1f} m</h2>
            </div>
        """, unsafe_allow_html=True)
    with c3: 
        st.markdown(f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">📐 TOTAL ACUMULADO</span>
                <h2 style="margin: 5px 0 0 0; color: #818cf8;">{metros_acumulados:.1f} m</h2>
            </div>
        """, unsafe_allow_html=True)
    with c4: 
        st.markdown(f"""
            <div class="card-destaque">
                <span style="font-size: 13px; color: #94a3b8;">⚡ EFICIÊNCIA OPERACIONAL</span>
                <h2 style="margin: 5px 0 0 0; color: #f43f5e;">{eficiencia:.0f}%</h2>
            </div>
        """, unsafe_allow_html=True)

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

            # Busca metragem total produzida por essa sonda específica
            if not df_prod.empty:
                prod_sonda = df_prod[df_prod["sonda_id"] == row["id"]]
                m_sonda = (prod_sonda["prof_final"] - prod_sonda["prof_inicial"]).sum() if not prod_sonda.empty else 0.0
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
    botao_voltar_dashboard()
    st.title("Gestão Central de Sondas")
    st.markdown("---")

    df_sondas = table_df("sondas")

    tab_lista, tab_novo = st.tabs(["Sondas Cadastradas", "Nova Sonda"])

    with tab_lista:
        if not df_sondas.empty:
            st.dataframe(df_sondas, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("🗑️ Excluir Sonda")
            
            opcoes_sonda = [f"{row['id']} - {row['codigo']} ({row['equipe']})" for _, row in df_sondas.iterrows()]
            sonda_excluir_str = st.selectbox("Selecione a Sonda para excluir:", opcoes_sonda)
            sonda_id_excluir = int(sonda_excluir_str.split(" - ")[0])

            if st.button("Excluir Sonda Selecionada", type="primary"):
                delete_row("sondas", "id", sonda_id_excluir)
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
                    try:
                        insert_row("sondas", {
                            "codigo": codigo,
                            "equipe": equipe,
                            "projeto": projeto,
                            "status": status,
                        })
                        st.success("Sonda cadastrada com sucesso!")
                        st.rerun()
                    except Exception as e:
                        if "duplicate" in str(e).lower() or "unique" in str(e).lower():
                            st.error("Sonda já cadastrada.")
                        else:
                            st.error(f"Erro ao cadastrar sonda: {e}")

# ------------------------------------------------------------------------------
# 3. APONTAMENTO DIÁRIO
# ------------------------------------------------------------------------------
elif opcao == "Apontamento Diário":
    botao_voltar_dashboard()
    st.title("Apontamento Diário de Produção")
    st.markdown("---")

    df_sondas = table_df("sondas")[["id", "codigo"]] if not table_df("sondas").empty else pd.DataFrame(columns=["id", "codigo"])
    df_furos = table_df("furos")
    df_prod = table_df("producao_diaria")
    if df_furos.empty:
        df_furos = pd.DataFrame(columns=["id", "sonda_id"])
    if df_prod.empty:
        df_prod = pd.DataFrame(columns=["id", "data", "sonda_id", "furo_id", "prof_inicial", "prof_final", "horas_trabalhadas", "horas_paradas", "motivo_parada"])
    if perfil_atual != "Admin":
        df_sondas = df_sondas[df_sondas["id"] == sonda_id_atual]
        df_furos = df_furos[df_furos["sonda_id"] == sonda_id_atual]
        df_prod = df_prod[df_prod["sonda_id"] == sonda_id_atual]
    if not df_prod.empty:
        df_prod = df_prod.merge(df_sondas[["id", "codigo"]], left_on="sonda_id", right_on="id", how="left", suffixes=("", "_sonda"))
        df_prod.rename(columns={"codigo": "sonda"}, inplace=True)
        df_prod["avanco"] = df_prod["prof_final"] - df_prod["prof_inicial"]
        df_prod = df_prod.sort_values("data", ascending=False)

    tab_hist, tab_novo = st.tabs(["Histórico", "Registrar Apontamento"])

    with tab_hist:
        if not df_prod.empty:
            st.dataframe(df_prod, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("🗑️ Excluir Registro de Apontamento")
            
            opcoes_prod = [f"ID: {row['id']} | Data: {row['data']} | Furo: {row['furo_id']} | Avanço: {row['avanco']}m" for _, row in df_prod.iterrows()]
            prod_excluir_str = st.selectbox("Selecione o Apontamento para excluir:", opcoes_prod)
            prod_id_excluir = int(prod_excluir_str.split(" | ")[0].replace("ID: ", ""))

            if st.button("Excluir Apontamento Selecionado", type="primary"):
                delete_row("producao_diaria", "id", prod_id_excluir)
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
                p_ini = c4.number_input("Prof. Inicial (m)", min_value=0.0, step=0.1)
                p_fin = c5.number_input("Prof. Final (m)", min_value=0.0, step=0.1)
                h_trab = c6.number_input("Horas Trabalhadas", min_value=0.0, step=0.5)

                c7, c8 = st.columns(2)
                h_par = c7.number_input("Horas Paradas", min_value=0.0, step=0.5)
                motivo = c8.text_input("Motivo Parada")

                if st.form_submit_button("Salvar Apontamento", type="primary"):
                    s_id = df_sondas[df_sondas["codigo"] == sonda_sel]["id"].values[0]
                    insert_row("producao_diaria", {
                        "data": str(dt),
                        "sonda_id": int(s_id),
                        "furo_id": furo_sel,
                        "prof_inicial": p_ini,
                        "prof_final": p_fin,
                        "horas_trabalhadas": h_trab,
                        "horas_paradas": h_par,
                        "motivo_parada": motivo,
                    })
                    st.success("Dados de produção salvos com sucesso!")
                    st.rerun()
        else:
            st.warning("É necessário possuir furos e sondas vinculados para registrar o apontamento.")

# ------------------------------------------------------------------------------
# 4. CONTROLE DE FUROS (COM CAPTURA DE GPS AUTOMÁTICO)
# ------------------------------------------------------------------------------
elif opcao == "Controle de Furos":
    botao_voltar_dashboard()
    st.title("Controle de Furos de Sondagem")
    st.markdown("---")

    df_sondas = table_df("sondas")
    df_furos = table_df("furos")
    if df_sondas.empty:
        df_sondas = pd.DataFrame(columns=["id", "codigo"])
    if df_furos.empty:
        df_furos = pd.DataFrame(columns=["id", "sonda_id", "coord_e", "coord_n", "cota", "prof_planejada", "prof_executada", "situacao"])
    if perfil_atual != "Admin":
        df_sondas = df_sondas[df_sondas["id"] == sonda_id_atual]
        df_furos = df_furos[df_furos["sonda_id"] == sonda_id_atual]
    if not df_furos.empty and not df_sondas.empty:
        df_furos = df_furos.merge(df_sondas[["id", "codigo"]], left_on="sonda_id", right_on="id", how="left", suffixes=("", "_sonda"))
        df_furos.rename(columns={"codigo": "sonda_codigo"}, inplace=True)

    tab_furos_list, tab_furos_novo = st.tabs(["Furos Cadastrados", "Novo Furo"])

    with tab_furos_list:
        if not df_furos.empty:
            st.dataframe(df_furos, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("🗑️ Excluir Furo")
            
            opcoes_furo = df_furos["id"].tolist()
            furo_excluir = st.selectbox("Selecione o Furo para excluir:", opcoes_furo)

            if st.button("Excluir Furo Selecionado", type="primary"):
                delete_row("furos", "id", furo_excluir)
                st.success(f"Furo '{furo_excluir}' excluído com sucesso!")
                st.rerun()
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
                    try:
                        insert_row("furos", {
                            "id": furo_id,
                            "sonda_id": int(s_id),
                            "coord_e": coord_e,
                            "coord_n": coord_n,
                            "cota": cota,
                            "prof_planejada": prof_plan,
                            "situacao": situacao,
                        })
                        st.success("Furo adicionado com sucesso!")
                        st.rerun()
                    except Exception as e:
                        if "duplicate" in str(e).lower() or "unique" in str(e).lower():
                            st.error("Identificação do Furo já existe.")
                        else:
                            st.error(f"Erro ao cadastrar furo: {e}")

# ------------------------------------------------------------------------------
# 5. BOLETIM GEOLÓGICO
# ------------------------------------------------------------------------------
elif opcao == "Boletim Geológico":
    botao_voltar_dashboard()
    st.title("Boletim Geológico de Campo")
    st.markdown("---")

    df_furos = table_df("furos")
    df_geo = table_df("boletim_geologico")
    if df_furos.empty:
        df_furos = pd.DataFrame(columns=["id", "sonda_id"])
    if df_geo.empty:
        df_geo = pd.DataFrame(columns=["id", "furo_id", "de_m", "ate_m", "recuperacao_m", "rqd_m", "litologia", "descricao_geologica", "n_amostra", "observacoes", "foto_url"])
    if perfil_atual != "Admin":
        df_furos = df_furos[df_furos["sonda_id"] == sonda_id_atual]
        df_geo = df_geo[df_geo["furo_id"].isin(df_furos["id"].tolist())]
    if not df_geo.empty:
        df_geo = df_geo.sort_values(["furo_id", "de_m"])

    tab_bg_hist, tab_bg_novo = st.tabs(["Registros Geológicos", "Registrar Intervalo"])

    with tab_bg_hist:
        if not df_geo.empty:
            st.dataframe(df_geo, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("🗑️ Excluir Registro do Boletim")
            
            lista_ids = df_geo["id"].tolist()
            id_para_excluir = st.selectbox("Selecione o ID do registro que deseja remover:", lista_ids)
            
            dados_registro = df_geo[df_geo["id"] == id_para_excluir].iloc[0]
            st.caption(f"Furo: **{dados_registro['furo_id']}** | Trecho: **{dados_registro['de_m']}m - {dados_registro['ate_m']}m** | Litologia: **{dados_registro['litologia']}**")
            
            if st.button("Excluir Registro Selecionado", type="primary"):
                delete_row("boletim_geologico", "id", int(id_para_excluir))

                st.success(f"Registro ID {id_para_excluir} excluído com sucesso!")
                st.rerun()
        else:
            st.info("Nenhum boletim registrado.")

    with tab_bg_novo:
        if not df_furos.empty:
            lista_furos = df_furos["id"].tolist()
            
            furo_sel = st.selectbox("Selecione o Furo para Registro", lista_furos)

            ultimo_resultado = (
                get_supabase()
                .table("boletim_geologico")
                .select("ate_m")
                .eq("furo_id", furo_sel)
                .order("ate_m", desc=True)
                .limit(1)
                .execute()
            )
            ultimo_ate = (_rows(ultimo_resultado)[0].get("ate_m") if _rows(ultimo_resultado) else None)

            de_auto = float(ultimo_ate) if ultimo_ate is not None else 0.0

            st.subheader("Intervalo e Métricas de Avanço")
            col_de, col_ate, col_avanco = st.columns(3)
            
            de_m = col_de.number_input("De (m)", min_value=0.0, value=de_auto, step=0.1, key="de_m_input")
            ate_m = col_ate.number_input("Até (m)", min_value=de_m, value=max(de_m, de_auto), step=0.1, key="ate_m_input")
            
            avanco_m = round(ate_m - de_m, 2)
            col_avanco.metric("Avanço Automático (m)", f"{avanco_m:.2f} m")

            with st.form("form_bg", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                rec_m = c1.number_input("Recuperação (m)", min_value=0.0, max_value=float(avanco_m) if avanco_m > 0 else 999.0, step=0.01)
                
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
                        insert_row("boletim_geologico", {
                            "furo_id": furo_sel,
                            "de_m": de_m,
                            "ate_m": ate_m,
                            "recuperacao_m": rec_m,
                            "rqd_m": rqd_m,
                            "litologia": litologia,
                            "descricao_geologica": desc,
                            "n_amostra": n_amostra,
                            "observacoes": obs,
                            "foto_url": foto_url,
                        })

                        # O registro é salvo apenas uma vez no Supabase.
                        # Mantemos a função de upload da foto separada.

                        st.success("Boletim Geológico salvo com sucesso!")
                        st.rerun()
        else:
            st.warning("Nenhum furo cadastrado para registrar boletim.")

# ------------------------------------------------------------------------------
# 6. GESTÃO DE USUÁRIOS (EXCLUSIVO ADMIN)
# ------------------------------------------------------------------------------
elif opcao == "Gestão de Usuários" and perfil_atual == "Admin":
    botao_voltar_dashboard()
    st.title("Vinculação de Usuários e Sondas")
    st.markdown("---")

    df_sondas = table_df("sondas")
    df_users = table_df("usuarios")
    if df_sondas.empty:
        df_sondas = pd.DataFrame(columns=["id", "codigo"])
    if df_users.empty:
        df_users = pd.DataFrame(columns=["id", "usuario", "perfil", "sonda_id"])
    if not df_users.empty and not df_sondas.empty:
        df_users = df_users.merge(df_sondas[["id", "codigo"]], left_on="sonda_id", right_on="id", how="left", suffixes=("", "_sonda"))
        df_users.rename(columns={"codigo": "sonda_vinculada"}, inplace=True)

    tab_u_lista, tab_u_novo = st.tabs(["Usuários Cadastrados", "Novo Usuário"])

    with tab_u_lista:
        if not df_users.empty:
            st.dataframe(df_users, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("🗑️ Excluir Usuário")
            
            # Não permite excluir o usuário admin principal para evitar bloqueio
            users_filtrados = df_users[df_users["usuario"] != "admin"]
            
            if not users_filtrados.empty:
                opcoes_user = [f"ID: {row['id']} - {row['usuario']} ({row['perfil']})" for _, row in users_filtrados.iterrows()]
                user_excluir_str = st.selectbox("Selecione o Usuário para excluir:", opcoes_user)
                user_id_excluir = int(user_excluir_str.split(" - ")[0].replace("ID: ", ""))

                if st.button("Excluir Usuário Selecionado", type="primary"):
                    delete_row("usuarios", "id", user_id_excluir)
                    st.success("Usuário excluído com sucesso!")
                    st.rerun()
            else:
                st.caption("Apenas o usuário 'admin' principal está cadastrado (protegido contra exclusão).")
        else:
            st.info("Nenhum usuário cadastrado.")

    with tab_u_novo:
        with st.form("form_user", clear_on_submit=True):
            new_user = st.text_input("Nome do Usuário")
            new_pass = st.text_input("Senha", type="password")
            new_perfil = st.selectbox("Perfil", ["Geólogo", "Operador", "Admin"])
            
            sonda_opcoes = ["Nenhuma / Admin"] + (df_sondas["codigo"].tolist() if not df_sondas.empty else [])
            new_sonda = st.selectbox("Sonda Vinculada", sonda_opcoes)

            if st.form_submit_button("Cadastrar Usuário", type="primary"):
                if new_user and new_pass:
                    s_id = None
                    if new_sonda != "Nenhuma / Admin" and not df_sondas.empty:
                        s_id = int(df_sondas[df_sondas["codigo"] == new_sonda]["id"].values[0])

                    try:
                        insert_row("usuarios", {
                            "usuario": new_user,
                            "senha": hash_senha(new_pass),
                            "perfil": new_perfil,
                            "sonda_id": s_id,
                        })
                        st.success("Usuário cadastrado com sucesso!")
                        st.rerun()
                    except Exception as e:
                        if "duplicate" in str(e).lower() or "unique" in str(e).lower():
                            st.error("Usuário já existe.")
                        else:
                            st.error(f"Erro ao cadastrar usuário: {e}")
