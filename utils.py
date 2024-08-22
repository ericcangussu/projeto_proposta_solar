import streamlit as st

def gerar_chave_unica(prefixo):
    chave_unica = f"{prefixo}_{st.session_state['unique_key']}"
    st.session_state['unique_key'] += 1
    return chave_unica
