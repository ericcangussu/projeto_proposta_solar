import streamlit as st

def init_session_state():
    if 'unique_key' not in st.session_state:
        st.session_state['unique_key'] = 0
    if 'itens_configurados' not in st.session_state:
        st.session_state['itens_configurados'] = []
    if 'eventos_pagamento' not in st.session_state:
        st.session_state['eventos_pagamento'] = {}
    if 'comentarios' not in st.session_state:
        st.session_state['comentarios'] = []
