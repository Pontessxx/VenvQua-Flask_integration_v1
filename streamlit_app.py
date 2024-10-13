import streamlit as st
import pandas as pd
import pyodbc
import plotly.graph_objects as go
import plotly.express as px
import numpy as np


# Definir o modo wide como padrão
st.set_page_config(layout="wide")

# Definindo cores globais para os tipos de frequência
data_dict = {
    'OK': {'cor': '#494949', 'marker': 'circle'},
    'FALTA': {'cor': '#FF5733', 'marker': 'x'},
    'ATESTADO': {'cor': '#FFC300', 'marker': 'diamond'},
    'CURSO': {'cor': '#8E44AD', 'marker': 'star'},
    'FÉRIAS': {'cor': '#a5a5a5', 'marker': 'square'},
    'ALPHAVILLE':{'cor': '#5D578E', 'marker': 'square'},
}

# Configurando a string de conexão
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\Users\Henrique\Downloads\Controle.accdb;"
)

# Conectar ao banco de dados Access
conn = pyodbc.connect(conn_str)

# Fazendo SELECT nos sites da tabela Site_Empresa
query_sites = """
SELECT DISTINCT Site.Sites
FROM Site_Empresa
INNER JOIN Site ON Site_Empresa.id_Sites = Site.id_Site
"""
df_sites = pd.read_sql(query_sites, conn)

# ---- Menu Lateral ----
st.sidebar.header("Filtros de Navegação")

# Caixa de seleção para o site no menu lateral
site_selecionado = st.sidebar.selectbox("Selecione o Site:", df_sites['Sites'].unique())


# Obter o ID do site selecionado
def get_site_id(site_name):
    cursor = conn.cursor()
    cursor.execute("SELECT id_Site FROM Site WHERE Sites = ?", (site_name,))
    result = cursor.fetchone()
    return result[0] if result else None

site_id = get_site_id(site_selecionado)

# Obter as empresas associadas ao site selecionado
def get_empresas(site_id):
    cursor = conn.cursor()
    query = """
    SELECT Empresa.id_Empresa, Empresa.Empresas
    FROM Site_Empresa
    INNER JOIN Empresa ON Site_Empresa.id_Empresas = Empresa.id_Empresa
    WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.Ativo = True
    """
    cursor.execute(query, (site_id,))
    empresas = [(row[0], row[1]) for row in cursor.fetchall()]
    return empresas

empresas = get_empresas(site_id)
empresa_opcoes = [empresa[1] for empresa in empresas]

# Caixa de seleção para a empresa (filtrada pelo site selecionado)
empresa_selecionada = st.sidebar.selectbox("Selecione a Empresa:", empresa_opcoes)

# Obter o ID da empresa selecionada
def get_empresa_id(empresa_nome):
    for empresa in empresas:
        if empresa[1] == empresa_nome:
            return empresa[0]
    return None

empresa_id = get_empresa_id(empresa_selecionada)

# Obter o ID_SiteEmpresa com base no site e empresa selecionados
def get_siteempresa_id(site_id, empresa_id):
    cursor = conn.cursor()
    query = """SELECT id_SiteEmpresa FROM Site_Empresa WHERE id_Sites = ? AND id_Empresas = ? AND Ativo = True"""
    cursor.execute(query, (site_id, empresa_id))
    result = cursor.fetchone()
    return result[0] if result else None

siteempresa_id = get_siteempresa_id(site_id, empresa_id)

# ---- Consulta Final com o ID_SiteEmpresa ----
query = """
SELECT Nome.Nome, Presenca.Presenca, Controle.Data
FROM Presenca 
INNER JOIN (Nome 
INNER JOIN Controle ON Nome.id_Nomes = Controle.id_Nome) 
ON Presenca.id_Presenca = Controle.id_Presenca
WHERE Controle.id_SiteEmpresa = ?
"""

# Lendo os dados filtrados pelo id_SiteEmpresa
df = pd.read_sql(query, conn, params=[siteempresa_id])

# Fechar a conexão
conn.close()


# ---- Filtros Acima da Tabela ----
st.title("Projeto de Visualização")

# Organizar as caixas de seleção em uma linha com st.columns()
col1, col2, col3, col4 = st.columns(4)

# Caixa de seleção para filtrar o nome
with col1:
    nome_selecionado = st.multiselect("Selecione o Nome:", df['Nome'].unique())

# Caixa de seleção para filtrar o tipo de presença
with col2:
    presenca_selecionada = st.multiselect("Selecione a Presença:", df['Presenca'].unique())

# ---- Correção no Filtro de Data ----
# Aqui garantimos que a coluna 'Data' seja formatada corretamente
df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')  # Converte para formato datetime

# Obter o intervalo de datas no DataFrame para definir os limites de seleção
data_min = df['Data'].min().date()
data_max = df['Data'].max().date()

# Seletor de data usando st.date_input (com colunas)
with col3:
    data_inicio, data_fim = st.date_input(
        "Selecione o Período de Data:",
        value=[data_min, data_max],
        min_value=data_min,
        max_value=data_max
    )
df['Mês'] = df['Data'].dt.strftime('%B')
with col4:
    mes_selecionado = st.multiselect("Seleicone o Mês: ", df['Mês'].unique())
# Verifica se o intervalo de data foi selecionado corretamente
df_filtrado = df.copy()
if data_inicio and data_fim:
    df_filtrado = df_filtrado[(df_filtrado['Data'] >= pd.to_datetime(data_inicio)) & 
                              (df_filtrado['Data'] <= pd.to_datetime(data_fim))]

if mes_selecionado:
    df_filtrado = df_filtrado[df_filtrado['Mês'].isin(mes_selecionado)]
# Aplicando filtros de nome e presença
if nome_selecionado:
    df_filtrado = df_filtrado[df_filtrado['Nome'].isin(nome_selecionado)]

if presenca_selecionada:
    df_filtrado = df_filtrado[df_filtrado['Presenca'].isin(presenca_selecionada)]

# ---- Formatando a Data para dd/mm/yyyy ----
# Aplicar formatação personalizada para exibição da coluna 'Data'
df_filtrado['Data'] = df_filtrado['Data'].dt.strftime('%d/%m/%Y')

# ---- Exibindo o DataFrame, Gráfico de Pizza e Indicadores ----
# Organizar a visualização: DataFrame, Gráfico de Pizza e Indicadores
col_df, col_pizza, col_indicadores = st.columns([2, 2, 1])
# ---- Calcular Indicadores ----
if not df_filtrado.empty:
    # Calcular o número de dias úteis no período
    dias_uteis = np.busday_count(data_inicio, data_fim, weekmask='1111100')  # De segunda a sexta-feira

    # Contar quantos dias de férias, OK, falta, atestado e curso estão no DataFrame
    total_ferias = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'FÉRIAS'])
    total_ok = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'OK'])
    total_faltas = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'FALTA'])
    total_atestados = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'ATESTADO'])
    total_curso = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'CURSO'])
else:
    dias_uteis = total_ferias = total_ok = total_faltas = total_atestados = total_curso = 0
# Exibir o DataFrame na primeira coluna
with col_df:
    st.write(df_filtrado)

# Exibir o gráfico de pizza na segunda coluna
with col_pizza:
    if not df_filtrado.empty:
        # Agrupar por tipo de presença e contar a frequência
        df_presenca = df_filtrado.groupby('Presenca').size().reset_index(name='counts')

        # Transformar as presenças em maiúsculas para combinar com o dicionário
        df_presenca['Presenca'] = df_presenca['Presenca'].str.upper()

        # Criar o gráfico de pizza usando Plotly, aplicando cores do dicionário
        fig_pizza = px.pie(
            df_presenca, 
            values='counts', 
            names='Presenca', 
            title='Distribuição do Tipo de Presença',
            color='Presenca',  # Usar o campo 'Presenca' para as cores
            color_discrete_map={tipo: data_dict[tipo]['cor'] for tipo in df_presenca['Presenca']}  # Mapeamento de cores
        )

        # Exibir o gráfico de pizza
        st.plotly_chart(fig_pizza)
    else:
        st.write("Nenhum dado disponível para o gráfico de pizza.")

# Exibir indicadores na terceira coluna
with col_indicadores:
    if not df_filtrado.empty:
        # Indicador 1: Quantidade de Dias Registrados (dias únicos)
        dias_registrados = df_filtrado['Data'].nunique()
        st.metric(label="Dias Registrados", value=dias_registrados)

        # Indicador 2: Quantidade de OK
        total_ok = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'OK'])
        st.metric(label="Total OK", value=total_ok)

        # Indicador 3: Quantidade de Faltas
        total_faltas = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'FALTA'])
        st.metric(label="Total Faltas", value=total_faltas)

        # Indicador 4: Quantidade de Atestados
        total_atestados = len(df_filtrado[df_filtrado['Presenca'].str.upper() == 'ATESTADO'])
        st.metric(label="Total Atestados", value=total_atestados)

        # Indicador 1: Dias Úteis no Período (substituindo o Total de Registros)
        st.metric(label="Dias Úteis no Período", value=dias_uteis)
    else:
        st.write("Sem dados para exibir os indicadores.")

# ---- Exibir o Gráfico de Dispersão ----
st.markdown("---")  # Separador visual

if not df_filtrado.empty:
    fig_dispersao = go.Figure()

    for presenca, info in data_dict.items():
        df_tipo = df_filtrado[df_filtrado['Presenca'].str.upper() == presenca]
        if not df_tipo.empty:
            fig_dispersao.add_trace(go.Scatter(
                x=df_tipo['Data'],
                y=df_tipo['Nome'],
                mode='markers',
                marker=dict(color=info['cor'], symbol=info['marker'], size=10),
                name=presenca
            ))

    # Personalizando o gráfico de dispersão sem fundo colorido
    fig_dispersao.update_layout(
        title= f'Presença no período: {data_inicio} até {data_fim}',
        font=dict(color='black'),  # Define a cor da fonte para preto
        xaxis=dict(showgrid=False, gridcolor='lightgray'),  # Exibe o grid com cor clara
        yaxis=dict(showgrid=False, gridcolor='lightgray')   # Exibe o grid com cor clara
    )

    # Exibir o gráfico de dispersão abaixo do DataFrame e Gráfico de Pizza
    st.plotly_chart(fig_dispersao)
else:
    st.write("Nenhum dado disponível para o gráfico de dispersão.")




# ---- Exibir Indicadores Antes do Gráfico de Barras Empilhadas ----
st.markdown("---")  # Separador visual para os indicadores

# Organizar indicadores em uma linha
col_ind1, col_ind2, col_ind3, col_ind4, col_ind5 = st.columns(5)


with col_ind1:
    st.markdown(f"<h5 style='color:{data_dict['FÉRIAS']['cor']};'>Total de Férias</h5>", unsafe_allow_html=True)
    st.metric(label="", value=total_ferias)

with col_ind2:
    st.markdown(f"<h5 style='color:{data_dict['OK']['cor']};'>Total de OK</h5>", unsafe_allow_html=True)
    st.metric(label="", value=total_ok)

with col_ind3:
    st.markdown(f"<h5 style='color:{data_dict['FALTA']['cor']};'>Total de Faltas</h5>", unsafe_allow_html=True)
    st.metric(label="", value=total_faltas)

with col_ind4:
    st.markdown(f"<h5 style='color:{data_dict['ATESTADO']['cor']};'>Total de Atestados</h5>", unsafe_allow_html=True)
    st.metric(label="", value=total_atestados)

with col_ind5:
    st.markdown(f"<h5 style='color:{data_dict['CURSO']['cor']};'>Total de Cursos</h5>", unsafe_allow_html=True)
    st.metric(label="", value=total_curso)

# ---- Exibir o Gráfico de Barras Empilhadas ----
st.markdown("---")  # Separador visual para o gráfico de barras empilhadas

# Criando o gráfico de barras empilhadas com Plotly Express
if not df_filtrado.empty:
    # Converter a coluna 'Presenca' para maiúsculas para combinar com as chaves do data_dict
    df_filtrado['Presenca'] = df_filtrado['Presenca'].str.upper()

    # Agrupar os dados por nome e presença, e contar as ocorrências
    df_agrupado = df_filtrado.groupby(['Nome', 'Presenca']).size().reset_index(name='counts')

    # Criando o gráfico de barras empilhadas com Plotly Express
    fig_barras_empilhadas = px.bar(
        df_agrupado, 
        x='Nome', 
        y='counts', 
        color='Presenca', 
        text='counts',
        title="Gráfico de Barras Empilhadas de Presenças",
        color_discrete_map={tipo: data_dict[tipo]['cor'] for tipo in df_agrupado['Presenca'].unique()},
    )

    # Atualizando o layout para exibir os valores dentro das barras
    fig_barras_empilhadas.update_traces(texttemplate='%{text}', textposition='inside')

    # Exibir o gráfico de barras empilhadas
    st.plotly_chart(fig_barras_empilhadas)
else:
    st.write("Nenhum dado disponível para o gráfico de barras empilhadas.")
