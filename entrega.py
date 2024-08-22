import streamlit as st

def gerenciar_prazos_entrega():
    st.subheader('Condições de Entrega')
    prazo_desenhos_aprovacao = st.number_input('Desenhos para aprovação (dias):', min_value=0, format="%d")
    prazo_aprovacao_cliente = st.number_input('Prazo para aprovação dos desenhos pelo cliente (dias):', min_value=0, format="%d")

    prazos_itens = []
    for index, item in enumerate(st.session_state['itens_configurados'], start=1):
        if item['Descrição'].startswith("TRANSFORMADOR"):
            continue

        titulo_item = f"ITEM {index:02d}: {item['Descrição'].split(' ')[0]}"
        with st.expander(f"{titulo_item}", expanded=False):
            prazo_entrega = st.text_input(f"Prazo de entrega para {titulo_item} (dias):", key=f"prazo_{item['Item']}")
            prazos_itens.append((titulo_item, prazo_entrega))

    resumo_texto = "### Resumo Entrega\n"
    resumo_texto += f"- Desenhos para aprovação: {prazo_desenhos_aprovacao} dias\n"
    resumo_texto += f"- Prazo para aprovação dos desenhos pelo cliente: {prazo_aprovacao_cliente} dias\n\n"

    for titulo, prazo in prazos_itens:
        resumo_texto += f"- Prazo de entrega para {titulo}: {prazo} dias\n"

    st.markdown(resumo_texto)
