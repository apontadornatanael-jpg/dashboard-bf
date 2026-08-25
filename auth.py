import sqlite3
import hashlib
import streamlit as st

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def criar_tabela_usuarios():
    conn = sqlite3.connect("central_sondagem.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            senha TEXT NOT NULL,
            perfil TEXT CHECK(perfil IN ('Admin', 'Geólogo', 'Operador')) NOT NULL
        )
    ''')
    
    cursor.execute("SELECT COUNT(*) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            "INSERT INTO usuarios (usuario, senha, perfil) VALUES (?, ?, ?)",
            ("admin", hash_senha("admin123"), "Admin")
        )
    conn.commit()
    conn.close()

def verificar_login(usuario, senha):
    conn = sqlite3.connect("central_sondagem.db")
    cursor = conn.cursor()
    cursor.execute(
        "SELECT perfil FROM usuarios WHERE usuario = ? AND senha = ?",
        (usuario, hash_senha(senha))
    )
    resultado = cursor.fetchone()
    conn.close()
    return resultado[0] if resultado else None

def tela_login():
    criar_tabela_usuarios()
    
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False
        st.session_state["usuario"] = ""
        st.session_state["perfil"] = ""

    if not st.session_state["autenticado"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.subheader("🔐 Acesso ao Sistema")
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            
            if st.button("Entrar", type="primary", use_container_width=True):
                perfil = verificar_login(usuario, senha)
                if perfil:
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = usuario
                    st.session_state["perfil"] = perfil
                    st.success("Login realizado com sucesso!")
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
        return False
    return True

def botao_logout():
    st.sidebar.markdown(f"👤 **{st.session_state['usuario']}** ({st.session_state['perfil']})")
    if st.sidebar.button("Sair / Logout"):
        st.session_state["autenticado"] = False
        st.session_state["usuario"] = ""
        st.session_state["perfil"] = ""
        st.rerun()
