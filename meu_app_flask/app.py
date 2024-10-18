from flask import Flask, render_template, request, jsonify
import pandas as pd
import pyodbc
import json
import warnings
import plotly.graph_objs as go
import plotly

# import click
# import logging

# log = logging.getLogger('werkzeug')
# log.setLevel(logging.ERROR)

# def secho(text, file=None, nl=None, err=None, color=None, **styles):
#     pass

# def echo(text, file=None, nl=None, err=None, color=None, **styles):
#     pass

# click.echo = echo
# click.secho = secho

warnings.filterwarnings('ignore')

app = Flask(__name__) 

# Configuração da conexão com o banco de dados Access
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\\Users\\Henrique\\Downloads\\Controle.accdb;"
)
conn = pyodbc.connect(conn_str)

# Dicionário para os meses em português
meses_dict = {
    "January": "01", "February": "02", "March": "03", "April": "04",
    "May": "05", "June": "06", "July": "07", "August": "08",
    "September": "09", "October": "10", "November": "11", "December": "12"
}

# Dicionário de cores e marcadores para cada tipo de presença
color_marker_map = {
    'OK': {'cor': '#494949', 'marker': 'circle'},
    'FALTA': {'cor': '#FF5733', 'marker': 'x'},
    'ATESTADO': {'cor': '#FFC300', 'marker': 'diamond'},
    'CURSO': {'cor': '#8E44AD', 'marker': 'star'},
    'FÉRIAS': {'cor': '#a5a5a5', 'marker': 'square'},
}

@app.route("/", methods=["GET", "POST"])
def index():
    # Consultar sites
    query_sites = "SELECT DISTINCT Sites FROM Site"
    sites = pd.read_sql(query_sites, conn)['Sites'].tolist()

    # Captura os valores dos filtros
    selected_site = request.form.get("site")
    selected_empresa = request.form.get("empresa")
    selected_nomes = request.form.getlist("nomes")
    selected_meses = request.form.getlist("meses")
    selected_presenca = request.form.getlist("presenca")

    empresas = []
    if selected_site:
        empresas = get_empresas(get_site_id(selected_site))

    # Inicializa a tabela como vazia
    df = pd.DataFrame(columns=['Nome', 'Presenca', 'Data'])

    # Variáveis para os gráficos
    pie_chart_data = {}
    scatter_chart_data = {}

    # Executa a consulta SQL somente se site e empresa forem selecionados
    if selected_site and selected_empresa:
        try:
            query = """
            SELECT Nome.Nome, Presenca.Presenca, Controle.Data
            FROM (((Controle
            INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
            INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca)
            INNER JOIN Site_Empresa ON Controle.id_SiteEmpresa = Site_Empresa.id_SiteEmpresa)
            WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.id_Empresas = ?
            """
            cursor = conn.cursor()
            cursor.execute(query, (get_site_id(selected_site), get_empresa_id(selected_empresa, empresas)))
            rows = cursor.fetchall()

            # Verificar se há dados retornados
            if rows:
                df = pd.DataFrame([list(row) for row in rows], columns=['Nome', 'Presenca', 'Data'])
                
                # Converte a coluna Data para datetime
                df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y')

                # Aplicar filtros adicionais
                if selected_nomes:
                    df = df[df['Nome'].isin(selected_nomes)]
                if selected_presenca:
                    df = df[df['Presenca'].isin(selected_presenca)]
                if selected_meses:
                    selected_meses_numeric = [meses_dict[mes] for mes in selected_meses]
                    df = df[df['Data'].dt.strftime('%m').isin(selected_meses_numeric)]

                # Formatar a data para exibição
                df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')

                # Gráfico de dispersão
                fig_dispersao = go.Figure()

                for presenca, info in color_marker_map.items():
                    df_tipo = df[df['Presenca'].str.upper() == presenca]
                    if not df_tipo.empty:
                        fig_dispersao.add_trace(go.Scatter(
                            x=df_tipo['Data'],
                            y=df_tipo['Nome'],
                            mode='markers',
                            marker=dict(color=info['cor'], symbol=info['marker'], size=10),
                            name=presenca
                        ))

                # Customizando o layout do gráfico de dispersão
                fig_dispersao.update_layout(
                    title=f'Presença no período',
                    xaxis=dict(showgrid=False, gridcolor='lightgray'),
                    yaxis=dict(showgrid=False, gridcolor='lightgray'),
                    font=dict(color='#999999'),
                    plot_bgcolor='rgba(0,0,0,0)',  # Remover o fundo da área de plotagem
                    paper_bgcolor='rgba(0,0,0,0)',  # Remover o fundo ao redor do gráfico
                    hovermode='closest'  # Exibir o valor mais próximo ao passar o mouse
                )


                # Converte o gráfico de dispersão para JSON para renderizar no HTML
                scatter_chart_data = json.dumps(fig_dispersao, cls=plotly.utils.PlotlyJSONEncoder)

                # Gráfico de Pizza (usando Chart.js)
                df_presenca = df.groupby('Presenca').size().reset_index(name='counts')
                labels = df_presenca['Presenca'].tolist()  # Tipos de presença
                values = df_presenca['counts'].tolist()    # Contagens de cada presença

                pie_chart_data = json.dumps({
                    'labels': labels,
                    'values': values
                })
                
        except Exception as e:
            print(f"Erro ao consultar ou criar DataFrame: {e}")
    else:
        pie_chart_data = None
        scatter_chart_data = None

    return render_template(
        "index.html",
        sites=sites,
        empresas=[e[1] for e in empresas],
        nomes=pd.read_sql("SELECT DISTINCT Nome FROM Nome", conn)['Nome'].tolist(),
        meses=meses_dict.keys(),
        presencas=pd.read_sql("SELECT DISTINCT Presenca FROM Presenca", conn)['Presenca'].tolist(),
        selected_site=selected_site,
        selected_empresa=selected_empresa,
        selected_nomes=selected_nomes,
        selected_meses=selected_meses,
        selected_presenca=selected_presenca,
        data=df,
        pie_chart_data=pie_chart_data,
        scatter_chart_data=scatter_chart_data,
        color_marker_map=color_marker_map,
    )

def get_site_id(site_name):
    cursor = conn.cursor()
    cursor.execute("SELECT id_Site FROM Site WHERE Sites = ?", (site_name,))
    result = cursor.fetchone()
    return result[0] if result else None

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

def get_empresa_id(empresa_nome, empresas):
    for empresa in empresas:
        if empresa[1] == empresa_nome:
            return empresa[0]
    return None

if __name__ == "__main__":
    # print('Runing on http://127.0.0.1/5000')
    app.run(debug=True)