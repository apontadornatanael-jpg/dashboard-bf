import streamlit as st

def aplicar_estilo_customizado():
    st.markdown("""
        <style>
        .metric-card {
            background-color: #f8f9fa;
            border-left: 5px solid #1F4E79;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 2px 2px 8px rgba(0,0,0,0.05);
            margin-bottom: 10px;
        }
        .metric-card h4 {
            margin: 0;
            color: #595959;
            font-size: 0.85rem;
            text-transform: uppercase;
        }
        .metric-card h2 {
            margin: 5px 0 0 0;
            color: #1F4E79;
            font-size: 1.6rem;
            font-weight: bold;
        }
        </style>
    """, unsafe_allow_html=True)

def render_kpi_card(titulo, valor):
    st.markdown(f"""
        <div class="metric-card">
            <h4>{titulo}</h4>
            2>{valor}</h2>
        </div>
    """, unsafe_allow_html=True)
