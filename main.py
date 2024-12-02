import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import base64
from io import BytesIO

# Função para carregar e organizar os dados
def load_and_clean_data(file):
    # Verificar se o arquivo é do tipo Excel (.xlsx)
    if file.type != 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        st.error('Por favor, carregue um arquivo Excel (.xlsx)')
        return None

    # Carregar o arquivo Excel
    try:
        xl = pd.ExcelFile(file)
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {str(e)}")
        return None

    # Criar um dicionário de DataFrames para cada aba
    dfs = {}
    for sheet_name in xl.sheet_names:
        try:
            df = xl.parse(sheet_name)
            # Selecionar apenas as colunas desejadas e remover linhas duplicadas pelo 'RA'
            colunas_desejadas = ['RA', 'ALUNO', 'CALOURO/VETERANO', 'COD NÍVEL ENSINO', 'CURSO', 'CPF', 'DATA MATRÍCULA',
                                 'STATUS PLETIVO', 'TIPO INGRESSO']
            colunas_presentes = [col for col in colunas_desejadas if col in df.columns]
            if not all(col in colunas_presentes for col in colunas_desejadas):
                st.warning(f"A aba '{sheet_name}' não contém todas as colunas desejadas. Colunas encontradas: {colunas_presentes}")
            df = df[colunas_presentes].drop_duplicates(subset=['RA'])
            dfs[sheet_name] = df
        except Exception as e:
            st.error(f"Erro ao carregar os dados da aba '{sheet_name}': {str(e)}")
            return None

    return dfs

# Função para criar o dashboard com gráficos e tabela
def create_dashboard(df):
    # Filtrar por calouros ou veteranos
    filtro_tipo_aluno = st.selectbox("Filtrar por tipo de aluno", ['Todos', 'Calouros', 'Veteranos'])
    filtro_status_pletivo = st.selectbox("Filtrar por Status Pletivo",
                                         ['Todos', 'Matriculado', 'Trancado', 'Cancelado', 'Aguardando Pgto P1'])

    # Obter as colunas disponíveis para permitir a seleção do usuário
    colunas_disponiveis = df[next(iter(df))].columns  # Pegando as colunas do primeiro DataFrame do dicionário

    # Verificar se a coluna 'COD NÍVEL ENSINO' está presente nas colunas disponíveis
    if 'COD NÍVEL ENSINO' in colunas_disponiveis:
        cod_nivel_ensino_unique = df[next(iter(df))]['COD NÍVEL ENSINO'].unique()
        filtro_cod_nivel_ensino = st.selectbox("Filtrar por Código de Nível de Ensino",
                                               ['Todos'] + cod_nivel_ensino_unique.tolist())
    else:
        st.warning("Coluna 'COD NÍVEL ENSINO' não encontrada nos dados carregados.")
        return

    # Verificar se a coluna 'TIPO INGRESSO' está presente nas colunas disponíveis
    if 'TIPO INGRESSO' in colunas_disponiveis:
        tipo_ingresso_unique = df[next(iter(df))]['TIPO INGRESSO'].unique()
        filtro_tipo_ingresso = st.selectbox("Filtrar por Tipo de Ingresso",
                                            ['Todos'] + tipo_ingresso_unique.tolist())
    else:
        st.warning("Coluna 'TIPO INGRESSO' não encontrada nos dados carregados.")
        return

    for sheet_name, df_sheet in df.items():
        st.write(f"### Dados da aba '{sheet_name}'")

        # Aplicar filtros
        if filtro_tipo_aluno == 'Calouros':
            df_sheet = df_sheet[df_sheet['CALOURO/VETERANO'] == 'CALOURO']
        elif filtro_tipo_aluno == 'Veteranos':
            df_sheet = df_sheet[df_sheet['CALOURO/VETERANO'] == 'VETERANO']

        if filtro_status_pletivo != 'Todos':
            df_sheet = df_sheet[df_sheet['STATUS PLETIVO'] == filtro_status_pletivo]

        if 'COD NÍVEL ENSINO' in colunas_disponiveis and filtro_cod_nivel_ensino != 'Todos':
            df_sheet = df_sheet[df_sheet['COD NÍVEL ENSINO'] == filtro_cod_nivel_ensino]

        if 'TIPO INGRESSO' in colunas_disponiveis and filtro_tipo_ingresso != 'Todos':
            df_sheet = df_sheet[df_sheet['TIPO INGRESSO'] == filtro_tipo_ingresso]

        # Mostrar a tabela com as informações dos alunos
        st.write("### Tabela com Informações dos Alunos")
        st.dataframe(df_sheet)

        # Botão para exportar a tabela para Excel
        if st.button(f'Exportar Tabela para Excel ({sheet_name})'):
            exportar_excel(df_sheet, sheet_name)

        # Contagem de alunos por status pletivo
        status_counts = df_sheet['STATUS PLETIVO'].value_counts().reset_index()
        status_counts.columns = ['STATUS PLETIVO', 'Quantidade']

        # Mostrar os flash cards com a contagem de status pletivo
        st.write("### Status dos Alunos")
        for index, row in status_counts.iterrows():
            st.markdown(f"""
                <div style="padding: 20px; border-radius: 10px; background-color: #333; color: #fff; margin: 10px 0; text-align: center; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);">
                    <h3 style="margin: 0;">{row['STATUS PLETIVO']}</h3>
                    <p style="font-size: 24px; margin: 5px 0;">{row['Quantidade']} alunos</p>
                </div>
            """, unsafe_allow_html=True)

        # Quantidade total de alunos por curso
        total_alunos_curso = df_sheet.groupby('CURSO').size().reset_index(name='Quantidade de Alunos')
        st.write("### Quantidade Total de Alunos por Curso")
        st.bar_chart(total_alunos_curso.set_index('CURSO'))

        # Ranking dos top 10 cursos com mais alunos matriculados
        top_cursos = total_alunos_curso.sort_values(by='Quantidade de Alunos', ascending=False).head(10)
        st.write("### Ranking dos Top 10 Cursos com Mais Alunos Matriculados")
        fig, ax = plt.subplots()
        ax.barh(top_cursos['CURSO'], top_cursos['Quantidade de Alunos'], color='skyblue')
        ax.set_xlabel('Quantidade de Alunos')
        ax.set_title('Top 10 Cursos com Mais Alunos Matriculados')
        st.pyplot(fig)

        # Distribuição de Calouros e Veteranos
        st.write("### Distribuição de Calouros e Veteranos")
        fig, ax = plt.subplots()
        sns.countplot(data=df_sheet, x='CALOURO/VETERANO', palette='pastel', ax=ax)
        ax.set_xlabel('Tipo de Aluno')
        ax.set_ylabel('Quantidade')
        ax.set_title('Distribuição de Calouros e Veteranos')
        st.pyplot(fig)

        # Quantidade de alunos por COD NÍVEL ENSINO
        if 'COD NÍVEL ENSINO' in colunas_disponiveis and filtro_cod_nivel_ensino == 'Todos':
            st.write("### Quantidade de Alunos por Código de Nível de Ensino")
            cod_nivel_ensino_counts = df_sheet['COD NÍVEL ENSINO'].value_counts().reset_index()
            cod_nivel_ensino_counts.columns = ['COD NÍVEL ENSINO', 'Quantidade']
            st.bar_chart(cod_nivel_ensino_counts.set_index('COD NÍVEL ENSINO'))
        elif 'COD NÍVEL ENSINO' in colunas_disponiveis and filtro_cod_nivel_ensino != 'Todos':
            df_filtered = df_sheet[df_sheet['COD NÍVEL ENSINO'] == filtro_cod_nivel_ensino]
            st.write(f"### Quantidade de Alunos para o Código de Nível de Ensino: {filtro_cod_nivel_ensino}")
            st.write(f"Total de Alunos: {len(df_filtered)}")

# Função para exportar a tabela para Excel
def exportar_excel(df, sheet_name):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    writer.save()
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{sheet_name}.xlsx">Download do arquivo Excel</a>'
    st.markdown(href, unsafe_allow_html=True)

# Configuração do Streamlit com título
st.set_page_config(page_title='Dashboard de Alunos')

# Título e descrição
st.title('Gestão de Alunos')
st.write('Carregue o arquivo Excel (.xlsx) para visualizar os dados.')

# Carregar o arquivo Excel
uploaded_file = st.file_uploader("Carregar arquivo Excel", type="xlsx")

# Se um arquivo for carregado, processar e criar o dashboard
if uploaded_file is not None:
    dfs = load_and_clean_data(uploaded_file)
    if dfs is not None:
        create_dashboard(dfs)
