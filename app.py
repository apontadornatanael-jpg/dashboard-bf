import os
import sqlite3
import hashlib
import math
import pandas as pd
import streamlit as st
from datetime import date
from streamlit_geolocation import streamlit_geolocation
from supabase import create_client, Client
from ui_components import aplicar_estilo_customizado, render_kpi_card

# ------------------------------------------------------------------------------
# 1. CONFIGURAÇÃO DA PÁGINA E ESTILOS
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Gestão de Sondagem & Geologia",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

aplicar_estilo_customizado()

# ------------------------------------------------------------------------------
# 2. CONEXÃO COM SUPABASE E BANCO LOCAL (SQLITE)
# ------------------------------------------------------------------------------
SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

@st.cache_resource
def get_supabase_client() -> Client:
    if SUPABASE_URL and SUPABASE_KEY:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    return None

supabase = get_supabase_client()

DB_NAME = "sondagem_geologia.db"

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela Sondas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sondas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE NOT NULL,
            modelo TEXT,
            status TEXT DEFAULT 'Operacional'
        )
    """)

    # Tabela Furos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS furos (
            id TEXT PRIMARY KEY,
            sonda_id INTEGER,
            coord_e REAL,
            coord_n REAL,
            cota REAL,
            prof_planejada REAL,
            prof_executada REAL DEFAULT 0.0,
            situacao TEXT DEFAULT 'Planejado',
            FOREIGN KEY (sonda_id) REFERENCES sondas(id)
        )
    """)

    # Tabela Produção Diária (Apontamento)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS producao_diaria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT NOT NULL,
            sonda_id INTEGER,
            furo_id TEXT,
            prof_inicial REAL,
            prof_final REAL,
            horas_trabalhadas REAL,
            horas_paradas REAL,
            motivo_parada TEXT,
            FOREIGN KEY (sonda_id) REFERENCES sondas(id),
            FOREIGN KEY (furo_id) REFERENCES furos(id)
        )
    """)

    # Tabela Boletim Geológico
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS boletim_geologico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            furo_id TEXT,
            de_m REAL,
            ate_m REAL,
            recuperacao_m REAL,
            rqd_m REAL,
            litologia TEXT,
            descricao_geologica TEXT,
            n_amostra TEXT,
            observacoes TEXT,
            foto_url TEXT,
            FOREIGN KEY (furo_id) REFERENCES furos(id)
        )
    """)

    # Tabela Usuários
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            senha TEXT NOT NULL,
            perfil TEXT NOT NULL,
            sonda_id INTEGER,
            FOREIGN KEY (sonda_id) REFERENCES sondas(id)
        )
    """)

    # Usuário padrão Admin se não existir
    cursor.execute("SELECT COUNT(*) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        admin_hash = hashlib.sha256("admin123".encode()).hexdigest()
        cursor.execute(
            "INSERT INTO usuarios (usuario, senha, perfil) VALUES (?, ?, ?)",
            ("admin", admin_hash, "Admin"),
        )

    conn.commit()
    conn.close()

init_db()

# ------------------------------------------------------------------------------
# 3. FUNÇÕES UTILITÁRIAS & SUPABASE
# ------------------------------------------------------------------------------
def hash_senha(senha: str) -> str:
    return hashlib.sha256(senha.encode()).hexdigest()

def latlon_to_utm(lat: float, lon: float):
    """Converte Lat/Lon para UTM aproximado (WGS84)."""
    zone = int((lon + 180) / 6) + 1
    lon0 = (zone - 1) * 6 - 180 + 3
    lat_rad = math.radians(lat)
    lon_rad = math.radians(lon)
    lon0_rad = math.radians(lon0)

    easting = 500000 + 6366197.724 * math.cos(lat_rad) * (lon_rad - lon0_rad)
    northing = 10000000 + 6366197.724 * lat_rad if lat < 0 else 6366197.724 * lat_rad
    return round(easting, 2), round(northing, 2), zone

def upload_foto_supabase(file_obj, file_name: str) -> str:
    if not supabase:
        st.warning("Supabase não configurado. Foto salva apenas localmente (simulação).")
        return f"local/{file_name}"
    try:
        bucket = "testemunhos"
        content = file_obj.read()
        file_path = f"fotos/{date.today()}_{file_name}"
        supabase.storage.from_(bucket).upload(file_path, content, {"content-type": file_obj.type})
        public_url = supabase.storage.from_(bucket).get_public_url(file_path)
        return public_url
    except Exception as e:
        st.error(f"Erro ao enviar foto para Supabase Storage: {e}")
        return None

def salvar_boletim_supabase(dados: dict):
    if supabase:
        try:
            supabase.table("boletim_geologico").insert(dados).execute()
        except Exception as e:
            st.warning(f"Não foi possível sincronizar com o Supabase Postgres: {e}")

# ------------------------------------------------------------------------------
# 4. CONTROLE DE SESSÃO & LOGIN
# ------------------------------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "usuario" not in st.session_state:
    st.session_state["usuario"] = None
if "perfil" not in st.session_state:
    st.session_state["perfil"] = None
if "sonda_id" not in st.session_state:
    st.session_state["sonda_id"] = None

def login():
    st.markdown("## 🔑 Acesso ao Sistema")
    with st.form("form_login"):
        user_input = st.text_input("Usuário")
        pass_input = st.text_input("Senha", type="password")
        btn_login = st.form_submit_button("Entrar", type="primary", use_container_width=True)

        if btn_login:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute(
                "SELECT usuario, perfil, sonda_id FROM usuarios WHERE usuario = ? AND senha = ?",
                (user_input, hash_senha(pass_input)),
            )
            row = cursor.fetchone()
            conn.close()

            if row:
                st.session_state["logged_in"] = True
                st.session_state["usuario"] = row[0]
                st.session_state["perfil"] = row[1]
                st.session_state["sonda_id"] = row[2]
                st.success(f"Bem-vindo(a), {row[0]}!")
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")

if not st.session_state["logged_in"]:
    login()
    st.stop()

# ------------------------------------------------------------------------------
# 5. SIDEBAR & NAVEGAÇÃO
# ------------------------------------------------------------------------------
with st.sidebar:
    if os.path.exists("logo_empresa.png"):
        st.image("logo_empresa.png", use_container_width=True)
    
    st.title("⛏️ Sondagem App")
    st.markdown(f"**Usuário:** {st.session_state['usuario']} | **Perfil:** {st.session_state['perfil']}")
    st.markdown("---")

    opcoes_menu = [
        "📊 Dashboard Geral",
        "🚜 Cadastro de Sondas",
        "📝 Apontamento Diário",
        "📍 Controle de Furos",
        "⛏️ Boletim Geológico",
    ]
    if st.session_state["perfil"] == "Admin":
        opcoes_menu.append("👥 Gestão de Usuários")

    opcao = st.radio("Navegação", opcoes_menu)

    st.markdown("---")
    if st.button("🚪 Sair / Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state["usuario"] = None
        st.session_state["perfil"] = None
        st.session_state["sonda_id"] = None
        st.rerun()

usuario_atual = st.session_state["usuario"]
perfil_atual = st.session_state["perfil"]
sonda_id_atual = st.session_state["sonda_id"]

# ------------------------------------------------------------------------------
# TELA 1. DASHBOARD GERAL
# ------------------------------------------------------------------------------
if opcao == "📊 Dashboard Geral":
    st.title("📊 Dashboard Executivo de Sondagem")
    st.caption("Visão geral de produtividade, metros perfurados e disponibilidade física.")
    st.markdown("---")

    conn = get_connection()

    if perfil_atual == "Admin":
        df_prod = pd.read_sql_query("SELECT * FROM producao_diaria", conn)
        df_furos = pd.read_sql_query("SELECT * FROM furos", conn)
        df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
    else:
        df_prod = pd.read_sql_query(
            "SELECT * FROM producao_diaria WHERE sonda_id = ?", conn, params=(sonda_id_atual,)
        )
        df_furos = pd.read_sql_query(
            "SELECT * FROM furos WHERE sonda_id = ?", conn, params=(sonda_id_atual,)
        )
        df_sondas = pd.read_sql_query(
            "SELECT * FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,)
        )
    conn.close()

    # --- TRATAMENTO ANTI-BUG TYPEERROR ---
    # Converte explicitamente colunas para numéricas e trata valores nulos (NaN / None)
    total_metros = 0.0
    total_horas_trab = 0.0
    total_horas_para = 0.0
    furos_concluidos = 0

    if not df_prod.empty:
        df_prod["prof_final"] = pd.to_numeric(df_prod["prof_final"], errors="coerce").fillna(0.0)
        df_prod["prof_inicial"] = pd.to_numeric(df_prod["prof_inicial"], errors="coerce").fillna(0.0)
        df_prod["horas_trabalhadas"] = pd.to_numeric(df_prod["horas_trabalhadas"], errors="coerce").fillna(0.0)
        df_prod["horas_paradas"] = pd.to_numeric(df_prod["horas_paradas"], errors="coerce").fillna(0.0)

        total_metros = float((df_prod["prof_final"] - df_prod["prof_inicial"]).sum())
        total_horas_trab = float(df_prod["horas_trabalhadas"].sum())
        total_horas_para = float(df_prod["horas_paradas"].sum())

    if not df_furos.empty:
        furos_concluidos = int((df_furos["situacao"] == "Concluído").sum())

    # Renderização dos KPIs
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_kpi_card("Metros Perfurados", f"{total_metros:.1f} m", "📏")
    with c2:
        render_kpi_card("Horas Trabalhadas", f"{total_horas_trab:.1f} h", "⏱️")
    with c3:
        render_kpi_card("Horas Paradas", f"{total_horas_para:.1f} h", "⚠️")
    with c4:
        render_kpi_card("Furos Concluídos", f"{furos_concluidos}", "✅")

    st.markdown("---")
    if not df_prod.empty:
        col_graf1, col_graf2 = st.columns(2)
        with col_graf1:
            st.subheader("Avanço Diário (m)")
            df_prod["avanço"] = df_prod["prof_final"] - df_prod["prof_inicial"]
            df_diario = df_prod.groupby("data")["avanço"].sum().reset_index()
            st.bar_chart(df_diario.set_index("data"))
        with col_graf2:
            st.subheader("Distribuição de Horas")
            df_horas = pd.DataFrame({
                "Tipo": ["Trabalhadas", "Paradas"],
                "Horas": [total_horas_trab, total_horas_para]
            })
            st.bar_chart(df_horas.set_index("Tipo"))
    else:
        st.info("Nenhum dado de produção cadastrado para exibição dos gráficos.")

# ------------------------------------------------------------------------------
# TELA 2. CADASTRO DE SONDAS
# ------------------------------------------------------------------------------
elif opcao == "🚜 Cadastro de Sondas":
    st.title("🚜 Gestão de Sondas e Equipamentos")
    st.caption("Cadastro e manutenção das sondas de perfuração.")
    st.markdown("---")

    conn = get_connection()
    df_sondas = pd.read_sql_query("SELECT * FROM sondas", conn)
    conn.close()

    tab_lista, tab_nova = st.tabs(["📋 Sondas Cadastradas", "➕ Nova Sonda"])

    with tab_lista:
        if not df_sondas.empty:
            st.dataframe(df_sondas, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma sonda cadastrada.")

    with tab_nova:
        if perfil_atual != "Admin":
            st.error("Apenas usuários Administradores podem cadastrar novas sondas.")
        else:
            with st.container(border=True):
                with st.form("form_sonda", clear_on_submit=True):
                    st.subheader("Cadastrar Nova Sonda")
                    c1, c2 = st.columns(2)
                    codigo = c1.text_input("Código/Identificador da Sonda (ex: SND-01)")
                    modelo = c2.text_input("Modelo / Fabricante (ex: Mach 1200, HW-60)")
                    status = st.selectbox("Status Operacional", ["Operacional", "Manutenção", "Inativa"])

                    btn_salvar_sonda = st.form_submit_button("Salvar Sonda", type="primary", use_container_width=True)

                    if btn_salvar_sonda and codigo:
                        conn = get_connection()
                        cursor = conn.cursor()
                        try:
                            cursor.execute(
                                "INSERT INTO sondas (codigo, modelo, status) VALUES (?, ?, ?)",
                                (codigo, modelo, status),
                            )
                            conn.commit()
                            st.success(f"Sonda '{codigo}' cadastrada com sucesso!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Já existe uma sonda cadastrada com este código.")
                        finally:
                            conn.close()

# ------------------------------------------------------------------------------
# TELA 3. APONTAMENTO DIÁRIO
# ------------------------------------------------------------------------------
elif opcao == "📝 Apontamento Diário":
    st.title("📝 Apontamento Diário de Perfuração (DPR)")
    st.caption("Lançamento de produção por turno, profundidades executadas e paradas de sonda.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_prod_full = pd.read_sql_query(
            """
            SELECT p.id, p.data, s.codigo as sonda, p.furo_id, p.prof_inicial, p.prof_final,
                   (p.prof_final - p.prof_inicial) as avanco_m, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada
            FROM producao_diaria p
            LEFT JOIN sondas s ON p.sonda_id = s.id ORDER BY p.data DESC, p.id DESC
        """,
            conn,
        )
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas WHERE status = 'Operacional'", conn)
        df_furos = pd.read_sql_query("SELECT id, sonda_id FROM furos WHERE situacao != 'Concluído'", conn)
    else:
        df_prod_full = pd.read_sql_query(
            """
            SELECT p.id, p.data, s.codigo as sonda, p.furo_id, p.prof_inicial, p.prof_final,
                   (p.prof_final - p.prof_inicial) as avanco_m, p.horas_trabalhadas, p.horas_paradas, p.motivo_parada
            FROM producao_diaria p
            LEFT JOIN sondas s ON p.sonda_id = s.id
            WHERE p.sonda_id = ? ORDER BY p.data DESC, p.id DESC
        """,
            conn,
            params=(sonda_id_atual,),
        )
        df_sondas = pd.read_sql_query(
            "SELECT id, codigo FROM sondas WHERE id = ? AND status = 'Operacional'", conn, params=(sonda_id_atual,)
        )
        df_furos = pd.read_sql_query(
            "SELECT id, sonda_id FROM furos WHERE sonda_id = ? AND situacao != 'Concluído'", conn, params=(sonda_id_atual,)
        )
    conn.close()

    tab_registros, tab_novo_apontamento, tab_excluir = st.tabs(
        ["📋 Apontamentos Realizados", "➕ Novo Apontamento", "🗑️ Excluir Registros"]
    )

    with tab_registros:
        if not df_prod_full.empty:
            st.dataframe(df_prod_full, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum apontamento diário realizado.")

    with tab_novo_apontamento:
        if not df_sondas.empty and not df_furos.empty:
            with st.container(border=True):
                st.subheader("Registrar Apontamento Diário")
                with st.form("form_apontamento", clear_on_submit=True):
                    c1, c2, c3 = st.columns(3)
                    data_reg = c1.date_input("Data do Apontamento", value=date.today())

                    sonda_sel_codigo = c2.selectbox("Sonda", df_sondas["codigo"].tolist())
                    sonda_id_sel = int(df_sondas[df_sondas["codigo"] == sonda_sel_codigo]["id"].values[0])

                    furos_filtrados = df_furos[df_furos["sonda_id"] == sonda_id_sel]["id"].tolist()
                    furo_sel = c3.selectbox(
                        "Furo de Sondagem",
                        furos_filtrados if furos_filtrados else ["Nenhum furo encontrado"],
                    )

                    c4, c5 = st.columns(2)
                    prof_inicial = c4.number_input("Profundidade Inicial (m)", min_value=0.0, step=0.1)
                    prof_final = c5.number_input("Profundidade Final (m)", min_value=0.0, step=0.1)

                    c6, c7 = st.columns(2)
                    horas_trabalhadas = c6.number_input("Horas Trabalhadas (h)", min_value=0.0, max_value=24.0, value=8.0, step=0.5)
                    horas_paradas = c7.number_input("Horas Paradas (h)", min_value=0.0, max_value=24.0, value=0.0, step=0.5)

                    motivo_parada = st.text_area("Motivo da Parada (se houver)", placeholder="Ex: Manutenção preventiva na bomba d'água, chuva forte...")

                    btn_salvar_prod = st.form_submit_button("Registrar Produção", type="primary", use_container_width=True)

                    if btn_salvar_prod:
                        if prof_final < prof_inicial:
                            st.error("A profundidade final não pode ser menor que a inicial.")
                        elif furo_sel == "Nenhum furo encontrado":
                            st.error("Selecione um furo válido para vincular o apontamento.")
                        else:
                            conn = get_connection()
                            cursor = conn.cursor()
                            try:
                                cursor.execute(
                                    """
                                    INSERT INTO producao_diaria 
                                    (data, sonda_id, furo_id, prof_inicial, prof_final, horas_trabalhadas, horas_paradas, motivo_parada)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                                    """,
                                    (str(data_reg), sonda_id_sel, furo_sel, prof_inicial, prof_final, horas_trabalhadas, horas_paradas, motivo_parada),
                                )
                                cursor.execute(
                                    """
                                    UPDATE furos 
                                    SET prof_executada = MAX(prof_executada, ?),
                                        situacao = CASE WHEN situacao = 'Planejado' THEN 'Em Andamento' ELSE situacao END
                                    WHERE id = ?
                                    """,
                                    (prof_final, furo_sel),
                                )
                                conn.commit()
                                st.success("Apontamento diário registrado com sucesso!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao salvar apontamento: {e}")
                            finally:
                                conn.close()
        else:
            st.warning("Cadastre primeiro ao menos uma Sonda e um Furo vinculados para poder registrar a produção.")

    with tab_excluir:
        if not df_prod_full.empty:
            with st.container(border=True):
                st.subheader("Excluir Apontamento")
                id_para_excluir = st.selectbox("Selecione o ID do Apontamento", df_prod_full["id"].tolist())
                if st.button("🗑️ Confirmar Exclusão", type="primary"):
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM producao_diaria WHERE id = ?", (id_para_excluir,))
                    conn.commit()
                    conn.close()
                    st.success(f"Apontamento #{id_para_excluir} excluído.")
                    st.rerun()

# ------------------------------------------------------------------------------
# TELA 4. CONTROLE DE FUROS
# ------------------------------------------------------------------------------
elif opcao == "📍 Controle de Furos":
    st.title("📍 Controle e Planejamento de Furos")
    st.caption("Gestão de posições, cotas e metas de perfuração com apoio de GPS.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_furos = pd.read_sql_query(
            """
            SELECT f.id, s.codigo as sonda_codigo, f.coord_e, f.coord_n, f.cota, 
                   f.prof_planejada, f.prof_executada, f.situacao
            FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id
        """,
            conn,
        )
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
    else:
        df_furos = pd.read_sql_query(
            """
            SELECT f.id, s.codigo as sonda_codigo, f.coord_e, f.coord_n, f.cota, 
                   f.prof_planejada, f.prof_executada, f.situacao
            FROM furos f LEFT JOIN sondas s ON f.sonda_id = s.id
            WHERE f.sonda_id = ?
        """,
            conn,
            params=(sonda_id_atual,),
        )
        df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas WHERE id = ?", conn, params=(sonda_id_atual,))
    conn.close()

    tab_lista_furos, tab_novo_furo, tab_status_furo = st.tabs(
        ["📋 Furos Cadastrados", "➕ Novo Furo", "🔄 Atualizar Situação"]
    )

    with tab_lista_furos:
        if not df_furos.empty:
            st.dataframe(df_furos, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum furo cadastrado.")

    with tab_novo_furo:
        if not df_sondas.empty:
            with st.container(border=True):
                st.subheader("Cadastrar Novo Furo de Sondagem")

                st.markdown("##### 🛰️ Obter Coordenadas do GPS Local")
                location = streamlit_geolocation()
                lat_gps, lon_gps = None, None
                if location and location.get("latitude"):
                    lat_gps = location["latitude"]
                    lon_gps = location["longitude"]
                    e_calc, n_calc, zone_calc = latlon_to_utm(lat_gps, lon_gps)
                    st.success(
                        f"GPS Capturado: Lat {lat_gps:.5f}, Lon {lon_gps:.5f} ➡️ UTM (Fuso {zone_calc}S): E={e_calc}, N={n_calc}"
                    )
                else:
                    e_calc, n_calc = 0.0, 0.0

                with st.form("form_furo", clear_on_submit=True):
                    c1, c2 = st.columns(2)
                    furo_id = c1.text_input("Identificação do Furo (ex: F-01, FS-03)")
                    sonda_sel = c2.selectbox("Sonda Designada", df_sondas["codigo"].tolist())
                    sonda_id_sel = int(df_sondas[df_sondas["codigo"] == sonda_sel]["id"].values[0])

                    c3, c4, c5 = st.columns(3)
                    coord_e = c3.number_input("Coordenada Este (UTM E)", value=float(e_calc), format="%.2f")
                    coord_n = c4.number_input("Coordenada Norte (UTM N)", value=float(n_calc), format="%.2f")
                    cota = c5.number_input("Cota (m)", value=0.0, step=0.5)

                    c6, c7 = st.columns(2)
                    prof_planejada = c6.number_input("Profundidade Planejada (m)", min_value=0.0, value=50.0, step=1.0)
                    situacao = c7.selectbox("Situação Inicial", ["Planejado", "Em Andamento", "Concluído", "Cancelado"])

                    btn_salvar_furo = st.form_submit_button("Salvar Furo", type="primary", use_container_width=True)

                    if btn_salvar_furo and furo_id:
                        conn = get_connection()
                        cursor = conn.cursor()
                        try:
                            cursor.execute(
                                """
                                INSERT INTO furos (id, sonda_id, coord_e, coord_n, cota, prof_planejada, situacao)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                                """,
                                (furo_id, sonda_id_sel, coord_e, coord_n, cota, prof_planejada, situacao),
                            )
                            conn.commit()
                            st.success(f"Furo '{furo_id}' cadastrado com sucesso!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("Já existe um furo cadastrado com essa identificação ID.")
                        finally:
                            conn.close()
        else:
            st.warning("É necessário ter ao menos uma sonda cadastrada para adicionar furos.")

    with tab_status_furo:
        if not df_furos.empty:
            with st.container(border=True):
                with st.form("form_status_furo"):
                    st.subheader("Alterar Situação do Furo")
                    furo_alterar = st.selectbox("Selecione o Furo", df_furos["id"].tolist())
                    nova_situacao = st.selectbox("Nova Situação", ["Planejado", "Em Andamento", "Concluído", "Cancelado"])
                    btn_atualizar_sit = st.form_submit_button("Atualizar Status", type="primary", use_container_width=True)

                    if btn_atualizar_sit:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute("UPDATE furos SET situacao = ? WHERE id = ?", (nova_situacao, furo_alterar))
                        conn.commit()
                        conn.close()
                        st.success(f"Situação do furo '{furo_alterar}' alterada para '{nova_situacao}'.")
                        st.rerun()

# ------------------------------------------------------------------------------
# TELA 5. BOLETIM GEOLÓGICO
# ------------------------------------------------------------------------------
elif opcao == "⛏️ Boletim Geológico":
    st.title("⛏️ Boletim Geológico & Descrição de Testemunhos")
    st.caption("Lançamento de litologia, porcentagem de recuperação, RQD e fotos no Supabase Storage.")
    st.markdown("---")

    conn = get_connection()
    if perfil_atual == "Admin":
        df_furos = pd.read_sql_query("SELECT id FROM furos", conn)
        df_geo = pd.read_sql_query(
            """
            SELECT bg.id, bg.furo_id, bg.de_m, bg.ate_m, (bg.ate_m - bg.de_m) as avanco_m,
                   bg.recuperacao_m, ROUND((bg.recuperacao_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as rec_pct,
                   bg.rqd_m, ROUND((bg.rqd_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as rqd_pct,
                   bg.litologia, bg.n_amostra, bg.descricao_geologica, bg.foto_url
            FROM boletim_geologico bg ORDER BY bg.furo_id, bg.de_m ASC
        """,
            conn,
        )
    else:
        df_furos = pd.read_sql_query("SELECT id FROM furos WHERE sonda_id = ?", conn, params=(sonda_id_atual,))
        df_geo = pd.read_sql_query(
            """
            SELECT bg.id, bg.furo_id, bg.de_m, bg.ate_m, (bg.ate_m - bg.de_m) as avanco_m,
                   bg.recuperacao_m, ROUND((bg.recuperacao_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as rec_pct,
                   bg.rqd_m, ROUND((bg.rqd_m / NULLIF(bg.ate_m - bg.de_m, 0)) * 100, 1) as rqd_pct,
                   bg.litologia, bg.n_amostra, bg.descricao_geologica, bg.foto_url
            FROM boletim_geologico bg JOIN furos f ON bg.furo_id = f.id 
            WHERE f.sonda_id = ? ORDER BY bg.furo_id, bg.de_m ASC
        """,
            conn,
            params=(sonda_id_atual,),
        )
    conn.close()

    tab_visu_geo, tab_novo_geo = st.tabs(["📋 Registros Descritos", "➕ Novo Inset / Manobra"])

    with tab_visu_geo:
        if not df_geo.empty:
            furo_filtro = st.selectbox("Filtrar por Furo", ["Todos"] + df_geo["furo_id"].unique().tolist())
            df_exibir = df_geo if furo_filtro == "Todos" else df_geo[df_geo["furo_id"] == furo_filtro]
            st.dataframe(df_exibir, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum boletim geológico registrado.")

    with tab_novo_geo:
        if not df_furos.empty:
            with st.container(border=True):
                st.subheader("Registrar Intervalo Geológico / Manobra")
                furo_geo_sel = st.selectbox("Furo de Sondagem", df_furos["id"].tolist())

                c1, c2 = st.columns(2)
                de_m = c1.number_input("De (m)", min_value=0.0, step=0.1)
                ate_m = c2.number_input("Até (m)", min_value=0.0, step=0.1)
                manobra = max(ate_m - de_m, 0.0)

                c3, c4 = st.columns(2)
                recuperacao_m = c3.number_input(
                    f"Recuperação (m) [Máx: {manobra:.2f}m]",
                    min_value=0.0,
                    max_value=float(manobra) if manobra > 0 else 0.0,
                    step=0.01,
                )
                rqd_m = c4.number_input(
                    f"RQD Sumatório (m) [Máx: {manobra:.2f}m]",
                    min_value=0.0,
                    max_value=float(manobra) if manobra > 0 else 0.0,
                    step=0.01,
                )

                rec_pct = (recuperacao_m / manobra * 100) if manobra > 0 else 0.0
                rqd_pct = (rqd_m / manobra * 100) if manobra > 0 else 0.0
                st.caption(f"📊 **Recuperação:** {rec_pct:.1f}% | **RQD:** {rqd_pct:.1f}%")

                c5, c6 = st.columns(2)
                litologia = c5.text_input("Litologia / Rocha (ex: Gnaisse, Basalto, Solo Argiloso)")
                n_amostra = c6.text_input("Nº da Amostra (opcional)")

                descricao_geologica = st.text_area("Descrição Geológico-Geotécnica (Grau de alteração, fraturamento, RMR, etc.)")
                observacoes = st.text_input("Observações Adicionais")

                st.markdown("##### 📸 Foto da Caixa de Testemunho")
                foto_upload = st.file_uploader("Selecione a Imagem", type=["jpg", "jpeg", "png"])

                if st.button("Salvar Boletim Geológico", type="primary", use_container_width=True):
                    if ate_m <= de_m:
                        st.error("O valor 'Até (m)' deve ser estritamente maior que 'De (m)'.")
                    elif not litologia:
                        st.error("Informe a litologia correspondente ao intervalo.")
                    else:
                        foto_url_final = None
                        if foto_upload is not None:
                            with st.spinner("Enviando foto para o Supabase Storage..."):
                                foto_url_final = upload_foto_supabase(foto_upload, foto_upload.name)

                        dados_boletim_local = {
                            "furo_id": furo_geo_sel,
                            "de_m": de_m,
                            "ate_m": ate_m,
                            "recuperacao_m": recuperacao_m,
                            "rqd_m": rqd_m,
                            "litologia": litologia,
                            "descricao_geologica": descricao_geologica,
                            "n_amostra": n_amostra,
                            "observacoes": observacoes,
                            "foto_url": foto_url_final,
                        }

                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            """
                            INSERT INTO boletim_geologico 
                            (furo_id, de_m, ate_m, recuperacao_m, rqd_m, litologia, descricao_geologica, n_amostra, observacoes, foto_url)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                furo_geo_sel,
                                de_m,
                                ate_m,
                                recuperacao_m,
                                rqd_m,
                                litologia,
                                descricao_geologica,
                                n_amostra,
                                observacoes,
                                foto_url_final,
                            ),
                        )
                        conn.commit()
                        conn.close()

                        salvar_boletim_supabase(dados_boletim_local)

                        st.success("Boletim Geológico gravado localmente e sincronizado no Supabase!")
                        st.rerun()
        else:
            st.warning("É preciso cadastrar ao menos um furo antes de lançar dados geológicos.")

# ------------------------------------------------------------------------------
# TELA 6. GESTÃO DE USUÁRIOS (APENAS ADMIN)
# ------------------------------------------------------------------------------
elif opcao == "👥 Gestão de Usuários":
    if perfil_atual != "Admin":
        st.error("Acesso restrito apenas a Administradores.")
        st.stop()

    st.title("👥 Gestão de Usuários & Permissões (RBAC)")
    st.caption("Controle de perfis (Admin, Geólogo, Operador) e vinculação às sondas.")
    st.markdown("---")

    conn = get_connection()
    df_users = pd.read_sql_query(
        """
        SELECT u.id, u.usuario, u.perfil, s.codigo as sonda_vinculada 
        FROM usuarios u LEFT JOIN sondas s ON u.sonda_id = s.id
    """,
        conn,
    )
    df_sondas = pd.read_sql_query("SELECT id, codigo FROM sondas", conn)
    conn.close()

    tab_users_lista, tab_user_novo = st.tabs(["📋 Usuários Cadastrados", "➕ Novo Usuário"])

    with tab_users_lista:
        st.dataframe(df_users, use_container_width=True, hide_index=True)

    with tab_user_novo:
        with st.container(border=True):
            with st.form("form_novo_usuario", clear_on_submit=True):
                st.subheader("Criar Novo Acesso")
                c1, c2 = st.columns(2)
                novo_user = c1.text_input("Nome de Usuário (Login)")
                nova_senha = c2.text_input("Senha", type="password")

                c3, c4 = st.columns(2)
                novo_perfil = c3.selectbox("Perfil de Acesso", ["Admin", "Geólogo", "Operador"])

                sonda_opcoes = ["Nenhuma (Acesso Global)"] + df_sondas["codigo"].tolist()
                sonda_vinc = c4.selectbox("Vincular Sonda", sonda_opcoes)

                btn_criar_user = st.form_submit_button("Criar Usuário", type="primary", use_container_width=True)

                if btn_criar_user and novo_user and nova_senha:
                    s_id = None
                    if sonda_vinc != "Nenhuma (Acesso Global)":
                        s_id = int(df_sondas[df_sondas["codigo"] == sonda_vinc]["id"].values[0])

                    conn = get_connection()
                    cursor = conn.cursor()
                    try:
                        cursor.execute(
                            "INSERT INTO usuarios (usuario, senha, perfil, sonda_id) VALUES (?, ?, ?, ?)",
                            (novo_user, hash_senha(nova_senha), novo_perfil, s_id),
                        )
                        conn.commit()
                        st.success(f"Usuário '{novo_user}' criado com sucesso!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Este nome de usuário já está em uso.")
                    finally:
                        conn.close()
