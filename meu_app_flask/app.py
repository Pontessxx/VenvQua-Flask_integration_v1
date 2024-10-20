from flask import Flask, render_template, request, jsonify, flash, redirect, url_for
import pandas as pd
import pyodbc
import json
import warnings
import plotly.graph_objs as go
import plotly
from datetime import datetime
import os
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
app.secret_key = "testeunique"
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

    # Variáveis para os gráficos (inicializando com None)
    pie_chart_data = None
    scatter_chart_data = None
    stacked_bar_chart_data = None
    total_dias_registrados = 0
    total_ok = 0
    total_faltas = 0
    total_atestados = 0
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
                    title={
                        'text': "Gráfico de disperssão de Presenças",
                        'x': 0.5,  # Centraliza o título
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': {'size': 24}  # Altera o tamanho da fonte do título
                    },
                    xaxis=dict(showgrid=False, gridcolor='lightgray'),
                    yaxis=dict(showgrid=False, gridcolor='lightgray'),
                    font=dict(color='#000000'),  # Cor padrão (será alterada via JavaScript)
                    plot_bgcolor='rgba(0,0,0,0)',  # Remover o fundo da área de plotagem
                    paper_bgcolor='rgba(0,0,0,0)',  # Remover o fundo ao redor do gráfico
                    hovermode='closest'
                )

                # Converte o gráfico de dispersão para JSON para renderizar no HTML
                scatter_chart_data = json.dumps(fig_dispersao, cls=plotly.utils.PlotlyJSONEncoder)

                # Gráfico de Pizza (usando Plotly)
                df_presenca = df.groupby('Presenca').size().reset_index(name='counts')
                labels = df_presenca['Presenca'].str.upper().tolist()  # Tipos de presença em maiúsculas
                values = df_presenca['counts'].tolist()    # Contagens de cada presença

                # Mapeamento das cores para o gráfico de pizza
                colors = [color_marker_map[label]['cor'] if label in color_marker_map else '#999999' for label in labels]

                # Criação do gráfico de pizza com Plotly
                fig_pie = go.Figure(data=[go.Pie(labels=labels, values=values, textinfo='label+percent', hole=0.3, marker=dict(colors=colors))])

                # Definir layout do gráfico de pizza
                fig_pie.update_layout(
                    title={
                        'text': "Distribuição de Presença",
                        'x': 0.5,  # Centraliza o título
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': {'size': 24}  # Altera o tamanho da fonte do título
                    },
                    showlegend=True,
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)'
                )

                # Converte o gráfico de pizza para JSON
                pie_chart_data = json.dumps(fig_pie, cls=plotly.utils.PlotlyJSONEncoder)

                df['Presenca'] = df['Presenca'].str.upper()
                df_agrupado = df.groupby(['Nome', 'Presenca']).size().reset_index(name='counts')
                barras = []

                for presenca in df_agrupado['Presenca'].unique():
                    df_presenca = df_agrupado[df_agrupado['Presenca'] == presenca]
                    barra = go.Bar(
                        x=df_presenca['Nome'],
                        y=df_presenca['counts'],
                        name=presenca,
                        marker=dict(color=color_marker_map[presenca]['cor']),
                        text=df_presenca['counts'],
                        textposition='inside'
                    )
                    barras.append(barra)

                layout = go.Layout(
                    title = {
                        'text': "Nomes x Presença",
                        'x': 0.5,  # Centraliza o título
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': {'size': 24}  # Altera o tamanho da fonte do título
                    },
                    barmode='stack',
                    xaxis=dict(title='Nome', showgrid=False),
                    yaxis=dict(title='Contagem de Presença', showgrid=False),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#000000')  # Cor padrão (será alterada via JavaScript)
                )

                fig_barras_empilhadas = go.Figure(data=barras, layout=layout)
                stacked_bar_chart_data = json.dumps(fig_barras_empilhadas, cls=plotly.utils.PlotlyJSONEncoder)
                
                total_dias_registrados = df['Data'].nunique()  # Contagem de dias únicos

                # Sum up the counts for OK, FALTAS, ATESTADO
                total_ok = df[df['Presenca'].str.upper() == 'OK'].shape[0]  # Contagem de OK
                total_faltas = df[df['Presenca'].str.upper() == 'FALTA'].shape[0]  # Contagem de FALTAS
                total_atestados = df[df['Presenca'].str.upper() == 'ATESTADO'].shape[0]  # Contagem de ATESTADOS

        except Exception as e:
            print(f"Erro ao consultar ou criar DataFrame: {e}")

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
        stacked_bar_chart_data=stacked_bar_chart_data,
        total_dias_registrados=total_dias_registrados,
        total_ok=total_ok,
        total_faltas=total_faltas,
        total_atestados=total_atestados,
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

def get_siteempresa_id(site_id, empresa_id):
    """Obtém o ID_SiteEmpresas com base no site e empresa selecionados, considerando apenas empresas ativas."""
    cursor = conn.cursor()
    query = """SELECT id_SiteEmpresa FROM Site_Empresa WHERE id_Sites = ? AND id_Empresas = ? AND Ativo = True"""
    cursor.execute(query, (site_id, empresa_id))
    result = cursor.fetchone()
    return result[0] if result else None

def get_nomes(siteempresa_id, ativos=True):
    """Obtém os nomes associados ao ID_SiteEmpresas, filtrando por ativos se solicitado."""
    cursor = conn.cursor()
    query = "SELECT Nome.Nome FROM Nome WHERE id_SiteEmpresa = ?"

    if ativos:
        query += " AND Ativo = True"
    else:
        query += " AND Ativo = False"

    cursor.execute(query, (siteempresa_id,))
    nomes = [row[0] for row in cursor.fetchall()]
    return nomes

@app.route('/adicionar-presenca', methods=['GET', 'POST'])
def adiciona_presenca():
    # Consultar sites e presenças
    query_sites = "SELECT DISTINCT Sites FROM Site"
    sites = pd.read_sql(query_sites, conn)['Sites'].tolist()
    presenca_opcoes = pd.read_sql("SELECT DISTINCT Presenca FROM Presenca", conn)['Presenca'].tolist()

    # Captura os valores dos filtros e converte para maiúsculas
    selected_site = request.form.get("site").upper() if request.form.get("site") else None
    selected_empresa = request.form.get("empresa").upper() if request.form.get("empresa") else None
    selected_nomes = [nome.upper() for nome in request.form.getlist("nomes")] if request.form.getlist("nomes") else []
    selected_presenca = request.form.get("presenca").upper() if request.form.get("presenca") else None
    
    siteempresa_id = None
    nomes = []
    nomes_desativados = []
    empresas = []
    
    # Obter ano e mês atuais
    current_year = datetime.now().year
    current_month = datetime.now().strftime("%m")  # Formato de dois dígitos para o mês
    
    # Gerar a lista de dias do mês
    dias = [str(i).zfill(2) for i in range(1, 32)]  # Gera a lista de dias de 01 a 31
    
    if selected_site:
        empresas = get_empresas(get_site_id(selected_site))
    
    if selected_site and selected_empresa:
        site_id = get_site_id(selected_site)
        empresa_id = get_empresa_id(selected_empresa, empresas)
        siteempresa_id = get_siteempresa_id(site_id, empresa_id)
        if siteempresa_id:
            nomes = get_nomes(siteempresa_id, ativos=True)
            nomes_desativados = get_nomes(siteempresa_id, ativos=False)  # Buscar nomes desativados
    
    # Renderiza o template HTML e passa as variáveis necessárias
    return render_template(
        "adicionar_presenca.html",
        sites=sites,
        empresas=[e[1] for e in empresas],
        selected_site=selected_site,
        selected_empresa=selected_empresa,
        siteempresa_id=siteempresa_id,
        nomes=nomes,  # Passa os nomes obtidos
        nomes_desativados=nomes_desativados,  # Passa os nomes desativados
        presenca_opcoes=presenca_opcoes,  # Passa as opções de presença
        dias=dias,  # Passa os dias do mês
        current_month=current_month,  # Passa o mês atual
        current_year=current_year,  # Passa o ano atual
        meses_dict=meses_dict,  # Dicionário de meses em português
        color_marker_map=color_marker_map,
    )


# __________________ ROTAS PARA FLUXO _________________
@app.route('/reativar-nome', methods=['POST'])
def reativar_nome():
    nome_desativado = request.form.get("nome_desativado").strip()
    siteempresa_id = request.form.get("siteempresa_id")  # Captura o siteempresa_id

    # Verifique os valores recebidos
    print(f"Nome desativado: {nome_desativado}, SiteEmpresa ID: {siteempresa_id}")

    if not nome_desativado:
        flash("Nenhum nome selecionado para reativar!", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        cursor = conn.cursor()
        cursor.execute("UPDATE Nome SET Ativo = True WHERE Nome = ? AND id_SiteEmpresa = ?",
                       (nome_desativado, siteempresa_id))
        conn.commit()
        flash(f"Nome {nome_desativado} reativado com sucesso!", "success")
    except Exception as e:
        print(f"Erro ao reativar nome: {e}")  # Saída para depuração
        flash(f"Erro ao reativar nome: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))


    
if __name__ == "__main__":
    # print('Runing on http://127.0.0.1/5000')
    app.run(debug=True)