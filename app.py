import streamlit as st
import pandas as pd
from io import BytesIO

def process_file(uploaded_file):
    # Lê o arquivo Excel
    df = pd.read_excel(uploaded_file)

    # Assume que os dados estão na primeira coluna
    col = df.columns[0]

    # Separa as medidas e a cor usando expressão regular atualizada
    df[['Medida1', 'Medida2', 'Medida3', 'Cor']] = df[col].str.extract(r'(\d+)[xX](\d+)[xX](\d+)\s+(.+)', expand=True)

    # Remove possíveis espaços nas colunas de medidas e cor
    df['Medida2'] = df['Medida2'].str.strip()
    df['Cor'] = df['Cor'].str.strip()

    # Verifica se a extração foi bem-sucedida
    if df[['Medida2', 'Cor']].isnull().any().any():
        st.error("Erro ao processar o arquivo. Verifique se o formato dos dados está correto.")
        return None

    # Agrupa pelos valores de Medida2 e Cor, e conta as ocorrências
    grouped_df = df.groupby(['Medida2', 'Cor']).size().reset_index(name='Quantidade')

    # Cria uma tabela dinâmica com Medida2 como linhas e Cor como colunas
    pivot_df = grouped_df.pivot_table(index='Medida2', columns='Cor', values='Quantidade', aggfunc='sum', fill_value=0)

    # Adiciona uma coluna 'Total' que soma as quantidades por linha (Medida2)
    pivot_df['Total'] = pivot_df.sum(axis=1)

    # Adiciona uma linha 'Total' que soma as quantidades por coluna (Cor)
    total_row = pd.DataFrame(pivot_df.sum(axis=0)).T
    total_row.index = ['Total']
    pivot_df = pd.concat([pivot_df, total_row])

    # Reseta o índice para que 'Medida2' seja uma coluna
    pivot_df = pivot_df.reset_index()

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

            # Botão para download do arquivo processado
            @st.cache_data
            def convert_df(df):
                # Converte o DataFrame para um objeto Excel em memória
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                processed_data = output.getvalue()
                return processed_data

            excel_bytes = convert_df(processed_df)

            st.download_button(
                label="Baixar arquivo processado em Excel",
                data=excel_bytes,
                file_name='salame.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

if __name__ == "__main__":
    main()