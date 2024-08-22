import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

# Caminho fixo para o arquivo de custos
CSV_PATH = 'tabela_custos_transformador.csv'

def carregar_csv():
    try:
        df = pd.read_csv(CSV_PATH)
        st.success(f"Arquivo {CSV_PATH} carregado com sucesso!")
        return df
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return None

def editar_tabela_custo(df):
    st.subheader("Editar Tabela de Custos")

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_pagination(paginationAutoPageSize=True)
    gb.configure_side_bar()
    gb.configure_default_column(editable=True, groupable=True)
    gb.configure_selection(selection_mode="multiple", use_checkbox=True)  # Permite a seleção de múltiplas linhas para exclusão

    grid_options = gb.build()

    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        enable_enterprise_modules=True,
        height=400,
        reload_data=True,
    )

    selected = grid_response['selected_rows']
    edited_df = pd.DataFrame(grid_response['data'])

    if st.button("Excluir Linhas Selecionadas"):
        if selected:  # Verifica se há linhas selecionadas
            try:
                indices = [int(row['_selectedRowNodeInfo']['nodeId']) for row in selected if '_selectedRowNodeInfo' in row]
                if indices:
                    df = df.drop(indices).reset_index(drop=True)
                    salvar_csv(df)
                    st.success("Linhas selecionadas excluídas com sucesso!")
                else:
                    st.warning("Nenhuma linha válida selecionada para exclusão.")
            except Exception as e:
                st.error(f"Erro ao excluir linhas: {e}")
        else:
            st.warning("Nenhuma linha selecionada para exclusão.")

    if st.button("Salvar Alterações"):
        salvar_csv(edited_df)
        st.success("Alterações salvas com sucesso!")

    return edited_df

def adicionar_nova_linha(df):
    st.subheader("Adicionar Nova Linha")

    nova_linha = {}
    for col in df.columns:
        if df[col].dtype == 'object':
            nova_linha[col] = st.text_input(f"Novo valor para {col}:")
        else:
            nova_linha[col] = st.number_input(f"Novo valor para {col}:", value=0)

    if st.button("Adicionar Linha"):
        nova_linha_df = pd.DataFrame([nova_linha])  # Cria um DataFrame com a nova linha
        df = pd.concat([df, nova_linha_df], ignore_index=True)  # Concatena ao DataFrame existente
        salvar_csv(df)
        st.success("Nova linha adicionada e salva com sucesso!")
    
    return df

def salvar_csv(df):
    df.to_csv(CSV_PATH, index=False)
    st.success(f"Tabela salva em {CSV_PATH}")

def main():
    st.title("Edição da Tabela de Custos - Transformador")
    
    # Carregar o CSV existente
    df = carregar_csv()

    if df is not None:
        # Adicionar nova linha
        df = adicionar_nova_linha(df)
        
        # Editar a tabela de custos
        df = editar_tabela_custo(df)

if __name__ == "__main__":
    main()