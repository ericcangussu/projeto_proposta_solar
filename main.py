import streamlit as st
from forms import dados_iniciais, escolha_tipo_proposta, configuracao_itens
from eventos import gerenciar_eventos_pagamento
from entrega import gerenciar_prazos_entrega
from doc_processing import gerar_documento
from config import init_session_state

# Inicializa os estados da sessão, se necessário
init_session_state()

st.title('Formulário de Proposta Solar')

# Seção: Dados Iniciais
dados_iniciais()

# Seção: Escolha do Tipo de Proposta
escolha_tipo_proposta()

# Seção: Configuração dos Itens
configuracao_itens()

# Seção: Eventos de Pagamento
gerenciar_eventos_pagamento()

# Seção: Condições de Entrega
gerenciar_prazos_entrega()

# Geração do Documento
if st.button('Gerar Documento com Tabela'):
    gerar_documento()
