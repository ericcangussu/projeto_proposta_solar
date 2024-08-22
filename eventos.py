import streamlit as st

def gerenciar_eventos_pagamento():
    st.subheader('Condições de Pagamento')

    for index, item in enumerate(st.session_state['itens_configurados'], start=1):
        descricao_item = item['Descrição'].split(' ')[0]

        # Definir o título resumido do item
        if descricao_item == "SKID":
            titulo_resumido = "SKID"
            sugestoes = [
                {"percentual": 25, "dias": 20, "referencia": "Aprovação de Desenhos"},
                {"percentual": 25, "dias": 30, "referencia": "Aprovação de Desenhos"},
                {"percentual": 25, "dias": 60, "referencia": "Aprovação de Desenhos"},
                {"percentual": 25, "dias": 0, "referencia": "Entrega do Equipamento"},
            ]
        elif descricao_item == "QGBT" or descricao_item == "TRANSFORMADOR":
            titulo_resumido = "QGBT e Transformador Isolado"
            sugestoes = [
                {"percentual": 25, "dias": 20, "referencia": "Aprovação de Desenhos"},
                {"percentual": 25, "dias": 30, "referencia": "Aprovação de Desenhos"},
                {"percentual": 25, "dias": 60, "referencia": "Aprovação de Desenhos"},
                {"percentual": 25, "dias": 0, "referencia": "Entrega do Equipamento"},
            ]
        elif descricao_item == "Cabine":
            titulo_resumido = "Cabine"
            sugestoes = [
                {"percentual": 20, "dias": 30, "referencia": "Aprovação de Desenhos"},
                {"percentual": 20, "dias": 60, "referencia": "Aprovação de Desenhos"},
                {"percentual": 20, "dias": 90, "referencia": "Aprovação de Desenhos"},
                {"percentual": 20, "dias": 120, "referencia": "Aprovação de Desenhos"},
                {"percentual": 20, "dias": 150, "referencia": "Aprovação de Desenhos"},
            ]
        else:  # Itens adicionais como Inversor, Módulo, etc.
            titulo_resumido = descricao_item
            sugestoes = [
                {"percentual": 100, "dias": 0, "referencia": "Aprovação de Desenhos"},
            ]

        titulo_item = f"ITEM {index:02d}: {titulo_resumido}"
        with st.expander(f"{titulo_item}", expanded=False):
            for i, sugestao in enumerate(sugestoes):
                with st.container():
                    col1, col2, col3, col4 = st.columns(4)
                    percentual = col1.number_input(
                        f"Percentual {i + 1}", min_value=0, max_value=100, value=sugestao["percentual"], key=f"percentual_{item['Item']}_{i}"
                    )
                    dias = col2.number_input(
                        f"Dias {i + 1}", min_value=0, value=sugestao["dias"], key=f"dias_{item['Item']}_{i}"
                    )
                    referencia = col3.selectbox(
                        f"Referência {i + 1}", ["Aprovação de Desenhos", "Entrega do Equipamento"], index=["Aprovação de Desenhos", "Entrega do Equipamento"].index(sugestao["referencia"]), key=f"referencia_{item['Item']}_{i}"
                    )
                    if col4.button(f"Excluir Linha {i + 1}", key=f"excluir_{item['Item']}_{i}"):
                        excluir_evento(index, i)

                    if item['Item'] not in st.session_state['eventos_pagamento']:
                        st.session_state['eventos_pagamento'][item['Item']] = []

                    if i >= len(st.session_state['eventos_pagamento'][item['Item']]):
                        st.session_state['eventos_pagamento'][item['Item']].append(
                            {"percentual": percentual, "dias": dias, "referencia": referencia}
                        )
                    else:
                        st.session_state['eventos_pagamento'][item['Item']][i] = {
                            "percentual": percentual, "dias": dias, "referencia": referencia
                        }

            st.write("**Eventos de Pagamento:**")
            for evento in st.session_state['eventos_pagamento'][item['Item']]:
                st.write(f"- {evento['percentual']}% - {evento['dias']} Dias após {evento['referencia']}")

def excluir_evento(item_index, evento_index):
    if item_index in range(len(st.session_state['itens_configurados'])):
        item = st.session_state['itens_configurados'][item_index]
        if evento_index in range(len(st.session_state['eventos_pagamento'][item['Item']])):
            del st.session_state['eventos_pagamento'][item['Item']][evento_index]
