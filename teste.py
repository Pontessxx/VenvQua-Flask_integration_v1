import dash
from dash import dcc, html, Input, Output
import pandas as pd
import pyodbc
import plotly.express as px
import numpy as np

# Configuração da conexão com Access
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\Users\Henrique\Downloads\Controle.accdb;"
)
conn = pyodbc.connect(conn_str)

# Query para buscar sites
query_sites = """
SELECT DISTINCT Site.Sites
FROM Site_Empresa
INNER JOIN Site ON Site_Empresa.id_Sites = Site.id_Site
"""
df_sites = pd.read_sql(query_sites, conn)

# Criação da aplicação Dash
app = dash.Dash(__name__)

# Layout do Dashboard
app.layout = html.Div([
    html.H1("Projeto de Visualização"),

    # Dropdown para seleção de site
    dcc.Dropdown(
        id='site-selecionado',
        options=[{'label': site, 'value': site} for site in df_sites['Sites'].unique()],
        value=df_sites['Sites'].iloc[0]
    ),
    
    # Dropdown para seleção de empresa (baseado no site)
    dcc.Dropdown(id='empresa-selecionada'),

    # Filtros de Nome e Presença
    dcc.Dropdown(id='nome-selecionado', multi=True),
    dcc.Dropdown(id='presenca-selecionada', multi=True),

    # Filtro de Data
    dcc.DatePickerRange(id='data-range'),

    # Gráfico de Pizza e de Dispersão
    dcc.Graph(id='grafico-pizza'),
    dcc.Graph(id='grafico-dispersao'),

    # Gráfico de Barras Empilhadas
    dcc.Graph(id='grafico-barras'),
])

# Callback para atualizar a lista de empresas com base no site selecionado
@app.callback(
    Output('empresa-selecionada', 'options'),
    Input('site-selecionado', 'value')
)
def update_empresas(site):
    cursor = conn.cursor()
    query = """
    SELECT Empresa.Empresas
    FROM Site_Empresa
    INNER JOIN Empresa ON Site_Empresa.id_Empresas = Empresa.id_Empresa
    WHERE Site.Sites = ?
    """
    cursor.execute(query, (site,))
    empresas = [row[0] for row in cursor.fetchall()]
    return [{'label': empresa, 'value': empresa} for empresa in empresas]

# Callback para atualizar gráficos e dados baseados nos filtros
@app.callback(
    [Output('grafico-pizza', 'figure'), Output('grafico-dispersao', 'figure'), Output('grafico-barras', 'figure')],
    [Input('empresa-selecionada', 'value'), Input('data-range', 'start_date'), Input('data-range', 'end_date')]
)
def update_graficos(empresa, data_inicio, data_fim):
    query = """
    SELECT Nome.Nome, Presenca.Presenca, Controle.Data
    FROM Presenca 
    INNER JOIN (Nome INNER JOIN Controle ON Nome.id_Nomes = Controle.id_Nome) 
    ON Presenca.id_Presenca = Controle.id_Presenca
    WHERE Controle.id_SiteEmpresa = ?
    """
    df = pd.read_sql(query, conn, params=[empresa])
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')

    # Filtros de data
    df = df[(df['Data'] >= data_inicio) & (df['Data'] <= data_fim)]

    # Gráfico de Pizza
    fig_pizza = px.pie(df, names='Presenca', title='Distribuição do Tipo de Presença')

    # Gráfico de Dispersão
    fig_dispersao = px.scatter(df, x='Data', y='Nome', color='Presenca', title='Presença ao longo do Tempo')

    # Gráfico de Barras Empilhadas
    df_grouped = df.groupby(['Nome', 'Presenca']).size().reset_index(name='counts')
    fig_barras = px.bar(df_grouped, x='Nome', y='counts', color='Presenca', text='counts', title='Presenças por Nome')

    return fig_pizza, fig_dispersao, fig_barras

# Executar a aplicação
if __name__ == '__main__':
    app.run_server(debug=True)
