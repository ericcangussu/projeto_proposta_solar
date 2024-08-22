import streamlit as st
import pandas as pd

def configurar_itens_proposta():
    if 'itens_configurados' not in st.session_state:
        st.session_state['itens_configurados'] = []

    def adicionar_item(descricao, qtde, valor):
        descricao_formatada = descricao.replace(" - ", "\n- ")
        total_valor = qtde * valor
        st.session_state.itens_configurados.append({
            "Item": len(st.session_state.itens_configurados) + 1,
            "Qtde": qtde,
            "Descrição": descricao_formatada,
            "Valor Unitário": valor,
            "Valor Total": total_valor
        })

    def limpar_tabela():
        st.session_state['itens_configurados'] = []

    st.subheader('Escolha do Tipo de Proposta')
    tipo_proposta = st.selectbox(
        'Selecione o tipo de proposta:',
        ['Selecione', 'Subestação Unitária', 'Estação Fotovoltaica']
    )

    if tipo_proposta != 'Selecione':
        st.subheader(f'Configuração da {tipo_proposta}')

        tipo_configuracao = st.selectbox(
            'Selecione a configuração:',
            ['SKID (Transformador Isolado + QBGT)', 'QGBT + Transformador Isolado']
        )

        if tipo_configuracao == 'SKID (Transformador Isolado + QBGT)':
            with st.expander('Configurar SKID'):
                potencia_skid = st.number_input('Potência do Transformador Isolado (kVA):', min_value=0, value=500)
                classe_tensao_skid = st.selectbox('Classe de Tensão do Transformador Isolado (kV):', ['15kV', '24kV', '36kV'])
                perdas_skid = st.text_input('Perdas (%):', value='1,2%')
                fator_k_skid = st.text_input('Fator K:', value='4')
                ip_skid = st.text_input('IP:', value='00')
                corrente_skid = st.number_input('Corrente do QBGT (A):', min_value=0, value=630)
                tensao_skid = st.number_input('Tensão do QBGT (V):', min_value=0, value=800)
                capacidade_curto_skid = st.number_input('Capacidade de Curto-Circuito (kA):', min_value=0, value=10)
                descricao_skid = f"SKID {potencia_skid} kVA\n- SKID ESTRUTURADO EM AÇO PARA {potencia_skid} kVA;\n- TRANSFORMADOR ISOLADO A SECO {potencia_skid} kVA {classe_tensao_skid} kV\n- PERDAS = {perdas_skid}; K = {fator_k_skid}; IP = {ip_skid};\n- QGBT {tensao_skid} V {corrente_skid} A {capacidade_curto_skid} kA"
                quantidade_skid = st.number_input('Quantidade de SKID:', min_value=1, value=1)
                valor_skid = st.number_input('Valor do SKID:', min_value=0.0, format="%.2f")
                if st.button('Adicionar SKID Configurado'):
                    adicionar_item(descricao_skid, quantidade_skid, valor_skid)

        elif tipo_configuracao == 'QGBT + Transformador Isolado':
            with st.expander('Configurar QGBT + Transformador Isolado', expanded=False):
                potencia_trafo = st.number_input('Potência (kVA):', min_value=0, value=500)
                classe_tensao_trafo = st.selectbox('Classe de Tensão:', ['15kV', '24kV', '36kV'])
                perdas_trafo = st.text_input('Perdas (%):', value='1,2%')
                fator_k_trafo = st.text_input('Fator K:', value='4')
                ip_trafo = st.text_input('IP:', value='00')
                descricao_trafo = f"TRANSFORMADOR ISOLADO A SECO {potencia_trafo} kVA {classe_tensao_trafo} kV - PERDAS = {perdas_trafo}; K = {fator_k_trafo}; IP = {ip_trafo};"
                quantidade_trafo = st.number_input('Quantidade de Transformadores:', min_value=1, value=1)
                valor_trafo = st.number_input('Valor do Transformador:', min_value=0.0, format="%.2f")
                corrente_qgbt = st.number_input('Corrente (A):', min_value=0, value=630)
                tensao_qgbt = st.number_input('Tensão (V):', min_value=0, value=800)
                capacidade_curto_qgbt = st.number_input('Capacidade de Curto-Circuito (kA):', min_value=0, value=10)
                descricao_qgbt = f"QGBT {tensao_qgbt}V {corrente_qgbt}A {capacidade_curto_qgbt}kA"
                quantidade_qgbt = st.number_input('Quantidade de QBGT:', min_value=1, value=1)
                valor_qgbt = st.number_input('Valor do QBGT:', min_value=0.0, format="%.2f")
                if st.button('Adicionar QGBT + Transformador Configurados'):
                    adicionar_item(descricao_trafo, quantidade_trafo, valor_trafo)
                    adicionar_item(descricao_qgbt, quantidade_qgbt, valor_qgbt)

            with st.expander('Configurar Cabine'):
                modelo_cabine = st.text_input('Concessionária:')
                quantidade_cabine = st.number_input('Quantidade:', min_value=1, value=1)
                valor_cabine = st.number_input('Valor:', min_value=0.0, format="%.2f")
                if st.button('Adicionar Cabine'):
                    descricao_cabine = f"Cabine {modelo_cabine}"
                    adicionar_item(descricao_cabine, quantidade_cabine, valor_cabine)

        if tipo_proposta == 'Estação Fotovoltaica':
            st.subheader('Configuração dos Itens Adicionais')

            with st.expander('Configurar Inversor'):
                marca_inversor = st.selectbox('Marca do Inversor', ['Canadian', 'GoodWe', 'Growatt', 'Huawei', 'Sungrow', 'Solis', 'Outras'])
                potencia_inversor = st.selectbox('Potência do Inversor (kW)', [75, 125, 200, 215, 250, 333, 350, 1100, 3125])
                acessorio_inversor = st.selectbox('Acessório do Inversor', ['COM100A', 'COM100E', 'PVS-16MH', 'Logger 3000A', 'Logger 3000B'])
                quantidade_inversor = st.number_input('Quantidade de Inversor:', min_value=1)
                valor_inversor = st.number_input('Valor do Inversor:', min_value=0.0, format="%.2f")
                if st.button('Adicionar Inversor Configurado'):
                    descricao_inversor = f"Inversor {marca_inversor} {potencia_inversor} kW com {acessorio_inversor}"
                    adicionar_item(descricao_inversor, quantidade_inversor, valor_inversor)

            with st.expander('Configurar Módulo'):
                marca_modulo = st.selectbox('Marca do Módulo', ['AE Solar', 'Canadian', 'JA Solar', 'Jinko', 'Longi', 'Risen', 'Trina', 'TW', 'Outras'])
                potencia_modulo = st.selectbox('Potência do Módulo (W)', [450, 455, 540, 545, 550, 555, 560, 565, 570, 575, 600, 650, 655, 660, 665, 680, 690, 695])
                quantidade_modulo = st.number_input('Quantidade de Módulo:', min_value=1)
                valor_modulo = st.number_input('Valor do Módulo:', min_value=0.0, format="%.2f")
                if st.button('Adicionar Módulo Configurado'):
                    descricao_modulo = f"Módulo {marca_modulo} {potencia_modulo} W"
                    adicionar_item(descricao_modulo, quantidade_modulo, valor_modulo)

            with st.expander('Configurar Cabine'):
                fabricante_cabine = st.selectbox('Fabricante da Cabine', ['Blutrafos', 'Gazquez', 'Lucy', 'Progressul', 'Setta', 'N/A'])
                modelo_medicao_cabine = st.selectbox('Modelo de Medição', ['1 Medição', '2 Medições', '3 Medições'])
                quantidade_cabine = st.number_input('Quantidade de Cabine:', min_value=1)
                valor_cabine = st.number_input('Valor da Cabine:', min_value=0.0, format="%.2f")
                if st.button('Adicionar Cabine Configurada'):
                    descricao_cabine = f"Cabine {fabricante_cabine} com {modelo_medicao_cabine}"
                    adicionar_item(descricao_cabine, quantidade_cabine, valor_cabine)

            with st.expander('Configurar Estrutura'):
                tipo_estrutura = st.selectbox('Tipo de Estrutura', ['Fixa', 'Tracker', 'Flutuante', 'N/A'])
                fixacao_tracker = st.selectbox('Fixação/Tracker', ['Axial', 'Brafer', 'Brametal', 'Dynamo', 'Isoeste', 'Metal Light', 'NTC Somar', 'SSM', 'Valmont', 'Outros', 'N/A'])
                quantidade_estrutura = st.number_input('Quantidade de Estrutura:', min_value=1)
                valor_estrutura = st.number_input('Valor da Estrutura:', min_value=0.0, format="%.2f")
                if st.button('Adicionar Estrutura Configurada'):
                    descricao_estrutura = f"Estrutura {tipo_estrutura} com fixação {fixacao_tracker}"
                    adicionar_item(descricao_estrutura, quantidade_estrutura, valor_estrutura)

            with st.expander('Configurar Cabo'):
                marca_cabo = st.selectbox('Marca do Cabo', ['Cabelauto', 'Conduspar', 'Halukabel', 'New Cabos', 'Nexans', 'Prysmian', 'Reicon', 'N/A'])
                tipo_cabo = st.selectbox('Tipo de Cabo', ['Média Tensão', 'Solar 4mm²', 'Solar 6mm²', 'Solar 8mm²', 'Solar 10mm²'])
                quantidade_cabo = st.number_input('Quantidade de Cabo:', min_value=1)
                valor_cabo = st.number_input('Valor do Cabo:', min_value=0.0, format="%.2f")
                if st.button('Adicionar Cabo Configurado'):
                    descricao_cabo = f"Cabo {marca_cabo} {tipo_cabo}"
                    adicionar_item(descricao_cabo, quantidade_cabo, valor_cabo)

    st.subheader('Tabela Resumo de Itens')

    if st.session_state.itens_configurados:
        df = pd.DataFrame(st.session_state.itens_configurados)
        df['Valor Unitário'] = df['Valor Unitário'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        df['Valor Total'] = df['Valor Total'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        total_geral = df['Valor Total'].apply(lambda x: float(x.replace("R$ ", "").replace(".", "").replace(",", "."))).sum()

        total_df = pd.DataFrame([["", "", "TOTAL", "", f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")]], columns=df.columns)
        df = pd.concat([df, total_df], ignore_index=True)

        st.table(df)
    else:
        st.write("Nenhum item configurado ainda.")

    if st.button('Limpar Tabela'):
        limpar_tabela()
