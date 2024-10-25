import streamlit as st
import pandas as pd
from io import StringIO

def process_file(uploaded_file):
    # Lê o arquivo Excel
    df = pd.read_excel(uploaded_file)

    # Assume que os dados estão na primeira coluna
    col = df.columns[0]

    # Converte a coluna para string
    df[col] = df[col].astype(str)

    # Separa as medidas e a cor usando expressão regular atualizada
    extract_result = df[col].str.extract(
        r'(\d+)[xX](\d+)[xX](\d+)\s+(.+)', expand=True
    )
    extract_result.columns = ['Medida1', 'Medida2', 'Medida3', 'Cor']

    # Atribui as colunas extraídas ao DataFrame original
    df = df.join(extract_result)

    # Verifica se a extração foi bem-sucedida
    if df[['Medida2', 'Cor']].isnull().any().any():
        st.error("Erro ao processar o arquivo. Verifique se o formato dos dados está correto.")
        return None

    # Remove possíveis espaços nas colunas de medidas e cor
    df['Medida2'] = df['Medida2'].str.strip()
    df['Cor'] = df['Cor'].str.strip()

    # Converte 'Medida2' para numérico para permitir ordenação
    df['Medida2_num'] = pd.to_numeric(df['Medida2'], errors='coerce')

    # Verifica novamente se há valores nulos após a conversão
    if df[['Medida2_num', 'Cor']].isnull().any().any():
        st.error("Erro ao processar o arquivo. Verifique se o formato dos dados está correto.")
        return None

    # Agrupa pelos valores de Medida2 e Cor, e conta as ocorrências
    grouped_df = df.groupby(['Medida2_num', 'Cor']).size().reset_index(name='Quantidade')

    # Cria uma tabela dinâmica com Medida2 como linhas e Cor como colunas
    pivot_df = grouped_df.pivot_table(
        index='Medida2_num',
        columns='Cor',
        values='Quantidade',
        aggfunc='sum',
        fill_value=0
    )

    # Adiciona uma coluna 'Total' que soma as quantidades por linha (Medida2)
    pivot_df['Total'] = pivot_df.sum(axis=1)

    # Converte o índice para coluna e renomeia para 'Medida2'
    pivot_df = pivot_df.reset_index().rename(columns={'Medida2_num': 'Medida2'})

    # Adiciona uma linha 'Total' que soma as quantidades por coluna (Cor)
    total_row = pivot_df.select_dtypes(include=[float, int]).sum(axis=0).to_frame().T
    total_row['Medida2'] = 'Total'
    pivot_df = pd.concat([pivot_df, total_row], ignore_index=True)

    # Ordena as medidas em ordem decrescente, mantendo 'Total' no final
    pivot_df['Medida2_sort'] = pd.to_numeric(pivot_df['Medida2'], errors='coerce')
    pivot_df = pivot_df.sort_values(
        by=['Medida2_sort', 'Medida2'],
        ascending=[False, False]
    ).drop(columns=['Medida2_sort'])

    # Ajusta a numeração das linhas
    pivot_df.index = range(1, len(pivot_df) + 1)
    pivot_df.index.name = 'Nº'

    # Cria sequência numérica para as colunas de cores
    cols = pivot_df.columns.tolist()
    color_columns = cols[1:-1]  # Exclui 'Medida2' e 'Total'
    sequence_numbers = [''] + [str(i+1) for i in range(len(color_columns))] + ['Total']

    # Achata o MultiIndex para exportação em CSV
    new_columns = []
    for seq, col in zip(sequence_numbers, cols):
        if seq:
            new_columns.append(f"{seq}-{col}")
        else:
            new_columns.append(col)

    pivot_df.columns = new_columns

    return pivot_df

def main():
    st.title("Processador de Arquivos Excel")

    st.write("Faça upload de um arquivo XLSX com as medidas e cores na mesma coluna.")

    # Upload do arquivo
    uploaded_file = st.file_uploader("Escolha um arquivo XLSX", type="xlsx")

    if uploaded_file is not None:
        st.success("Arquivo carregado com sucesso!")

        # Processa o arquivo
        processed_df = process_file(uploaded_file)

        if processed_df is not None:
            # Exibe os dados processados
            st.subheader("Dados Processados")
            st.dataframe(processed_df)

            # Converte o DataFrame para CSV
            @st.cache_data
            def convert_df_to_csv(df):
                return df.to_csv(index=True, sep=';', decimal=',', encoding='utf-8')

            csv_data = convert_df_to_csv(processed_df)

            # Botão de download do CSV
            st.download_button(
                label="Baixar arquivo processado em CSV",
                data=csv_data,
                file_name='salame.csv',
                mime='text/csv'
            )

if __name__ == "__main__":
    main()