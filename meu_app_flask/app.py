from flask import Flask, render_template, request, jsonify, flash, redirect, url_for, session
import pandas as pd
import pyodbc # type: ignore
import json
import warnings
import plotly.graph_objs as go
import plotly
from datetime import datetime, timedelta
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
    "Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
    "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
    "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"
}

# Dicionário de cores e marcadores para cada tipo de presença
color_marker_map = {
    'OK': {'cor': '#494949', 'marker': 'circle'},
    'FALTA': {'cor': '#FF5733', 'marker': 'x'},
    'ATESTADO': {'cor': '#FFC300', 'marker': 'diamond'},
    'FOLGA': {'cor': '#233F7B', 'marker': 'diamond'},
    'CURSO': {'cor': '#8E44AD', 'marker': 'star'},
    'FÉRIAS': {'cor': '#a5a5a5', 'marker': 'square'},
    'ALPHAVILLE':{'cor': '#76A9B7', 'marker': 'square'},
}

@app.route("/", methods=["GET", "POST"])
def index():
    # Consultar sites
    query_sites = "SELECT DISTINCT Sites FROM Site"
    sites = pd.read_sql(query_sites, conn)['Sites'].tolist()

    # Captura os valores dos filtros
    selected_site = request.form.get("site") or session.get('selected_site')
    selected_empresa = request.form.get("empresa") or session.get('selected_empresa')
    selected_ano = request.form.get("ano")  # Captura o valor do ano selecionado

    # Salva os valores na sessão
    if selected_site:
        session['selected_site'] = selected_site
    if selected_empresa:
        session['selected_empresa'] = selected_empresa

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
            query_params = [get_site_id(selected_site), get_empresa_id(selected_empresa, empresas)]
                
            # Filtro de ano
            if selected_ano:
                query += " AND YEAR(Controle.Data) = ?"
                query_params.append(selected_ano)

            cursor = conn.cursor()
            cursor.execute(query, query_params)
            rows = cursor.fetchall()

            # Verificar se há dados retornados
            if rows:
                df = pd.DataFrame([list(row) for row in rows], columns=['Nome', 'Presenca', 'Data'])

                # Converte a coluna Data para datetime
                df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m-%d %H:%M:%S')

                # Aplicar filtros adicionais
                if selected_nomes:
                    df = df[df['Nome'].isin(selected_nomes)]
                if selected_presenca:
                    df = df[df['Presenca'].isin(selected_presenca)]
                if selected_meses:
                    selected_meses_numeric = [meses_dict[mes] for mes in selected_meses]
                    df = df[df['Data'].dt.strftime('%m').isin(selected_meses_numeric)]

                # Gera uma lista contínua de datas entre o menor e o maior valor de data
                min_data = df['Data'].min()
                max_data = df['Data'].max()
                datas_continuas = pd.date_range(min_data, max_data).to_list()

                # Cria uma nova DataFrame com todas as combinações possíveis de nomes e datas contínuas
                nomes_unicos = df['Nome'].unique()
                df_continuo = pd.MultiIndex.from_product([nomes_unicos, datas_continuas], names=['Nome', 'Data']).to_frame(index=False)

                # Converte ambas as colunas 'Data' para datetime para garantir a compatibilidade no merge
                df_continuo['Data'] = pd.to_datetime(df_continuo['Data'])
                df['Data'] = pd.to_datetime(df['Data'])

                # Faz o merge do DataFrame original com o DataFrame contínuo
                df_merge = pd.merge(df_continuo, df, on=['Nome', 'Data'], how='left')

                # Preenche valores ausentes com "invisível" ou algum valor placeholder
                df_merge['Presenca'].fillna('invisível', inplace=True)

                # Gráfico de dispersão
                fig_dispersao = go.Figure()

                for presenca, info in color_marker_map.items():
                    df_tipo = df_merge[df_merge['Presenca'].str.upper() == presenca]
                    if not df_tipo.empty:
                        fig_dispersao.add_trace(go.Scatter(
                            x=df_tipo['Data'],
                            y=df_tipo['Nome'],
                            mode='markers',
                            marker=dict(color=info['cor'], symbol=info['marker'], size=10),
                            name=presenca
                        ))

                # Adicionar os pontos invisíveis para garantir o espaçamento correto
                df_invisivel = df_merge[df_merge['Presenca'] == 'invisível']
                fig_dispersao.add_trace(go.Scatter(
                    x=df_invisivel['Data'],
                    y=df_invisivel['Nome'],
                    mode='markers',
                    marker=dict(color='rgba(0,0,0,0)', size=10),  # Invisível
                    name='invisível',
                    showlegend=False  # Não mostrar na legenda
                ))

                # Customizando o layout do gráfico de dispersão
                fig_dispersao.update_layout(
                    title={
                        'text': "Gráfico de Dispersão de Presenças",
                        'x': 0.5,
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': {'size': 24}
                    },
                    xaxis=dict(
                        showgrid=False,
                        gridcolor='lightgray',
                        tickformat='%d/%m/%Y'  # Formata as datas no eixo X como dd/mm/yyyy
                    ),
                    yaxis=dict(showgrid=False, gridcolor='lightgray'),
                    font=dict(color='#000000'),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
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

                # Gráfico de Barras Empilhadas
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
                    title={
                        'text': "Nomes x Presença",
                        'x': 0.5,  # Centraliza o título
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': {'size': 24}  # Altera o tamanho da fonte do título
                    },
                    barmode='stack',
                    width=360,
                    xaxis=dict(title='Nome', showgrid=False),
                    yaxis=dict(title='Contagem de Presença', showgrid=False),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#000000')
                )

                fig_barras_empilhadas = go.Figure(data=barras, layout=layout)
                stacked_bar_chart_data = json.dumps(fig_barras_empilhadas, cls=plotly.utils.PlotlyJSONEncoder)

                # Contagem de dias únicos para o resumo
                total_dias_registrados = df['Data'].nunique()  # Contagem de dias únicos
                total_ok = df[df['Presenca'].str.upper() == 'OK'].shape[0]  # Contagem de OK
                total_faltas = df[df['Presenca'].str.upper() == 'FALTA'].shape[0]  # Contagem de FALTAS
                total_atestados = df[df['Presenca'].str.upper() == 'ATESTADO'].shape[0]  # Contagem de ATESTADOS

                # Formatar a coluna 'Data' para o formato 'dd/mm/yyyy' para a tabela
                df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')

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
        data=df,  # Agora com as datas formatadas para dd/mm/yyyy
        pie_chart_data=pie_chart_data,
        scatter_chart_data=scatter_chart_data,  # Gráfico de dispersão com datas formatadas
        stacked_bar_chart_data=stacked_bar_chart_data,
        total_dias_registrados=total_dias_registrados,
        total_ok=total_ok,
        total_faltas=total_faltas,
        total_atestados=total_atestados,
        color_marker_map=color_marker_map,
        selected_ano=selected_ano,
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


def get_empresas_inativas(site_id):
    cursor = conn.cursor()
    query = """
    SELECT Empresa.id_Empresa, Empresa.Empresas
    FROM Site_Empresa
    INNER JOIN Empresa ON Site_Empresa.id_Empresas = Empresa.id_Empresa
    WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.Ativo = False
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
    selected_site = request.form.get("site") or session.get('selected_site')
    selected_empresa = request.form.get("empresa") or session.get('selected_empresa')

    # Salva os valores na sessão
    if selected_site:
        session['selected_site'] = selected_site
    if selected_empresa:
        session['selected_empresa'] = selected_empresa

    selected_nomes = [nome.upper() for nome in request.form.getlist("nomes")] if request.form.getlist("nomes") else []
    selected_presenca = request.form.get("presenca").upper() if request.form.get("presenca") else None
    
    siteempresa_id = None
    nomes = []
    nomes_desativados = []
    empresas = []
    empresas_inativas = []  # Adicionando as inativas

    # Obter ano e mês atuais
    current_year = datetime.now().year
    current_month = datetime.now().strftime("%m")  # Formato de dois dígitos para o mês
    
    # Gerar a lista de dias do mês
    dias = [str(i).zfill(2) for i in range(1, 32)]  # Gera a lista de dias de 01 a 31
    
    if selected_site:
        empresas = get_empresas(get_site_id(selected_site))  # Empresas ativas
        empresas_inativas = get_empresas_inativas(get_site_id(selected_site))  # Empresas inativas

    registros_mes_atual = []

    if selected_site and selected_empresa:
        site_id = get_site_id(selected_site)
        empresa_id = get_empresa_id(selected_empresa, empresas)
        siteempresa_id = get_siteempresa_id(site_id, empresa_id)
        
        # Consulta para pegar os registros do mês e ano atual
        query = """
            SELECT Nome.Nome, Presenca.Presenca, Controle.Data
            FROM (((Controle
            INNER JOIN Nome ON Controle.id_Nome = Nome.id_Nomes)
            INNER JOIN Presenca ON Controle.id_Presenca = Presenca.id_Presenca)
            INNER JOIN Site_Empresa ON Controle.id_SiteEmpresa = Site_Empresa.id_SiteEmpresa)
            WHERE Site_Empresa.id_Sites = ? AND Site_Empresa.id_Empresas = ?
            AND MONTH(Controle.Data) = ? AND YEAR(Controle.Data) = ?
        """
        cursor = conn.cursor()
        cursor.execute(query, (site_id, empresa_id, current_month, current_year))
        registros_mes_atual = cursor.fetchall()  # Pega os registros

        if siteempresa_id:
            nomes = get_nomes(siteempresa_id, ativos=True)
            nomes_desativados = get_nomes(siteempresa_id, ativos=False)  # Buscar nomes desativados
    
    # Renderiza o template HTML e passa as variáveis necessárias
    return render_template(
        "adicionar_presenca.html",
        sites=sites,
        empresas=[e[1] for e in empresas],  # Empresas ativas
        empresas_inativas=[e[1] for e in empresas_inativas],  # Passa as inativas para o template
        selected_site=selected_site,
        selected_empresa=selected_empresa,
        siteempresa_id=siteempresa_id,
        nomes=nomes,  # Passa os nomes obtidos
        nomes_desativados=nomes_desativados,  # Passa os nomes desativados
        presenca_opcoes=presenca_opcoes,  # Passa as opções de presença
        dias=dias,  # Passa os dias do mês
        current_month=current_month,  # Passa o mês atual
        current_year=current_year,  # Passa o ano atual
        registros_mes_atual=registros_mes_atual,  # Passa os registros do mês atual
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

@app.route('/inativar-nome', methods=['POST'])
def inativar_nome():
    nome_ativo = request.form.get("nome_ativo").strip()
    siteempresa_id = request.form.get("siteempresa_id")  # Captura o siteempresa_id

    # Verifique os valores recebidos
    print(f"Nome ativo: {nome_ativo}, SiteEmpresa ID: {siteempresa_id}")

    if not nome_ativo:
        flash("Nenhum nome selecionado para desativar!", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        cursor = conn.cursor()

        # Verificar quantos nomes ativos existem para a SiteEmpresa selecionada
        cursor.execute("SELECT COUNT(*) FROM Nome WHERE id_SiteEmpresa = ? AND Ativo = True", (siteempresa_id,))
        num_nomes_ativos = cursor.fetchone()[0]

        # Impedir a desativação se houver apenas um nome ativo
        if num_nomes_ativos <= 1:
            flash("Não é possível desativar o último nome ativo. Pelo menos um nome deve permanecer ativo.", "error")
            return redirect(url_for('adiciona_presenca'))

        # Marcar o nome como inativo
        cursor.execute("UPDATE Nome SET Ativo = False WHERE Nome = ? AND id_SiteEmpresa = ?", (nome_ativo, siteempresa_id))
        conn.commit()

        flash(f"Nome {nome_ativo} desativado com sucesso!", "success")
    except Exception as e:
        print(f"Erro ao desativar nome: {e}")  # Saída para depuração
        flash(f"Erro ao desativar nome: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))


@app.route('/presenca', methods=['POST'])
def controlar_presenca():
    # Captura os valores dos filtros e dados do formulário
    nomes = request.form.getlist('nomes')  # Captura os nomes selecionados
    tipo_presenca = request.form.get('presenca')  # Captura o tipo de presença
    dia = request.form.get('dia')  # Captura o dia
    mes = request.form.get('mes')  # Captura o mês
    ano = request.form.get('ano')  # Captura o ano
    siteempresa_id = request.form.get('siteempresa_id')  # Captura o siteempresa_id
    action_type = request.form.get('action_type')  # Captura o tipo de ação (adicionar/remover)

    # Verificar se todos os dados foram fornecidos
    if not nomes or not dia or not mes or not ano:
        flash("Por favor, selecione todos os campos: Nomes, Dia, Mês e Ano.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        # Converte a data para datetime e verifica o dia da semana
        data_selecionada = datetime(int(ano), int(mes), int(dia))  # Converte a data para datetime
        dia_semana = data_selecionada.weekday()  # Retorna o dia da semana (0 = segunda-feira, 6 = domingo)

        # Impede a inserção de presença em sábados (5) e domingos (6)
        if dia_semana >= 5:
            flash("Não é permitido adicionar presença em sábados ou domingos.", "error")
            return redirect(url_for('adiciona_presenca'))

        cursor = conn.cursor()

        nomes_adicionados = []
        nomes_atualizados = []

        if action_type == 'adicionar':
            for nome in nomes:
                cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
                id_nome = cursor.fetchone()[0]

                # Verifica se já existe um registro para o nome e data
                cursor.execute("""
                    SELECT id_Controle FROM Controle 
                    WHERE id_Nome = ? AND Data = ? AND id_SiteEmpresa = ?
                """, (id_nome, data_selecionada, siteempresa_id))
                id_controle = cursor.fetchone()

                cursor.execute("SELECT id_Presenca FROM Presenca WHERE Presenca = ?", (tipo_presenca,))
                id_presenca = cursor.fetchone()[0]

                if id_controle:
                    # Se já existe um registro, atualize-o
                    cursor.execute("""
                        UPDATE Controle 
                        SET id_Presenca = ?
                        WHERE id_Controle = ?
                    """, (id_presenca, id_controle[0]))
                    nomes_atualizados.append(nome)
                else:
                    # Caso contrário, insira um novo registro
                    cursor.execute("""
                        INSERT INTO Controle (id_Nome, id_Presenca, Data, id_SiteEmpresa)
                        VALUES (?, ?, ?, ?)
                    """, (id_nome, id_presenca, data_selecionada, siteempresa_id))
                    nomes_adicionados.append(nome)

            conn.commit()  # Confirmar as alterações no banco de dados

            # Exibir mensagens separadas para nomes adicionados e atualizados
            if nomes_adicionados:
                flash(f"Presença adicionada com sucesso para os nomes: {', '.join(nomes_adicionados)} na data {data_selecionada.strftime('%d/%m/%Y')}", "success")
            if nomes_atualizados:
                flash(f"Presença atualizada com sucesso para os nomes: {', '.join(nomes_atualizados)} na data {data_selecionada.strftime('%d/%m/%Y')}", "warning")

        elif action_type == 'remover':
            # Remover presença
            for nome in nomes:
                cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
                id_nome = cursor.fetchone()[0]

                cursor.execute("""
                    SELECT id_Controle FROM Controle
                    WHERE id_Nome = ? AND Data = ? AND id_SiteEmpresa = ?
                """, (id_nome, data_selecionada, siteempresa_id))

                id_controle = cursor.fetchone()

                if id_controle:
                    cursor.execute("DELETE FROM Controle WHERE id_Controle = ?", (id_controle[0],))
                else:
                    flash(f"Não foi encontrado registro de presença para {nome} na data {data_selecionada.strftime('%d/%m/%Y')}.", "error")

            conn.commit()  # Confirmar as alterações no banco de dados
            flash(f"Presença removida para os nomes: {', '.join(nomes)} na data {data_selecionada.strftime('%d/%m/%Y')}", "remover")

    except pyodbc.Error as e:
        flash(f"Erro ao realizar a ação de presença: {e}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/adicionar-nome', methods=['POST'])
def adicionar_nome():
    novo_nome = request.form.get("novo_nome")
    siteempresa_id = request.form.get("siteempresa_id")

    if not novo_nome or not siteempresa_id:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        # Formatar o nome: primeira letra maiúscula, o restante em minúsculas
        novo_nome = novo_nome.strip().title()

        # Verificar se o nome já existe na tabela para o mesmo siteempresa_id
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (novo_nome, siteempresa_id))
        existe_nome = cursor.fetchone()[0]

        if existe_nome > 0:
            flash(f"O nome '{novo_nome}' já existe na tabela.", "warning")
            return redirect(url_for('adiciona_presenca'))

        # Pega o último id_Nomes e soma 1 para criar um novo ID
        cursor.execute("SELECT MAX(id_Nomes) FROM Nome")
        ultimo_id = cursor.fetchone()[0]
        novo_id = ultimo_id + 1

        # Insere o novo nome na tabela Nome
        cursor.execute("""
            INSERT INTO Nome (id_Nomes, id_SiteEmpresa, Nome, Ativo)
            VALUES (?, ?, ?, ?)
        """, (novo_id, siteempresa_id, novo_nome, True))

        conn.commit()
        flash(f"Nome '{novo_nome}' adicionado com sucesso!", "success")
    except Exception as e:
        flash(f"Erro ao adicionar nome: {e}", "error")

    return redirect(url_for('adiciona_presenca'))


@app.route('/adicionar-empresa', methods=['POST'])
def adicionar_empresa():
    site_nome = request.form.get("site") or session.get('selected_site')
    nova_empresa = request.form.get("nova_empresa").strip()

    # Verifica se os campos foram preenchidos
    if not site_nome or not nova_empresa:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        # Inserir a nova empresa na tabela Empresa
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM Empresa WHERE Empresas = ?", (nova_empresa,))
        existe_empresa = cursor.fetchone()[0]

        if existe_empresa > 0:
            flash(f"A empresa '{nova_empresa}' já existe.", "warning")
            return redirect(url_for('adiciona_presenca'))

        # Pega o último id_Empresa e soma 1 para criar um novo ID
        cursor.execute("SELECT MAX(id_Empresa) FROM Empresa")
        ultimo_id_empresa = cursor.fetchone()[0]
        novo_id_empresa = ultimo_id_empresa + 1

        # Inserir a nova empresa na tabela Empresa
        cursor.execute("""
            INSERT INTO Empresa (id_Empresa, Empresas)
            VALUES (?, ?)
        """, (novo_id_empresa, nova_empresa))

        # Pegar o ID do site selecionado
        site_id = get_site_id(site_nome)

        # Verifica se o site foi encontrado
        if not site_id:
            flash("Site não encontrado.", "error")
            return redirect(url_for('adiciona_presenca'))

        # Inserir a associação na tabela Site_Empresa
        cursor.execute("""
            INSERT INTO Site_Empresa (id_Sites, id_Empresas, Ativo)
            VALUES (?, ?, ?)
        """, (site_id, novo_id_empresa, True))

        conn.commit()  # Confirma as alterações no banco de dados
        flash(f"Empresa '{nova_empresa}' adicionada com sucesso ao site '{site_nome}'!", "success")
    except Exception as e:
        flash(f"Erro ao adicionar empresa: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))


@app.route('/desativar-empresa', methods=['POST'])
def desativar_empresa():
    empresa_ativa = request.form.get("empresa_ativa")  # Captura o nome da empresa ativa

    if not empresa_ativa:
        flash("Nenhuma empresa selecionada para desativar.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        # Verificar quantas empresas ativas existem
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM Site_Empresa WHERE Ativo = True")
        num_empresas_ativas = cursor.fetchone()[0]

        # Impedir a desativação da última empresa ativa
        if num_empresas_ativas <= 1:
            flash("Não é possível desativar todas as empresas. Pelo menos uma empresa deve estar ativa.", "error")
            return redirect(url_for('adiciona_presenca'))

        # Impedir a desativação da empresa atualmente selecionada na sessão
        empresa_selecionada = session.get('selected_empresa')
        if empresa_selecionada == empresa_ativa:
            flash(f"A empresa '{empresa_ativa}' está em uso e não pode ser desativada.", "error")
            return redirect(url_for('adiciona_presenca'))

        # Buscar o id da empresa selecionada
        cursor.execute("SELECT id_Empresa FROM Empresa WHERE Empresas = ?", (empresa_ativa,))
        id_empresa = cursor.fetchone()[0]

        # Atualizar o status da empresa na tabela Site_Empresa para inativa
        cursor.execute("UPDATE Site_Empresa SET Ativo = False WHERE id_Empresas = ?", (id_empresa,))
        conn.commit()

        flash(f"Empresa '{empresa_ativa}' desativada com sucesso!", "success")
    except Exception as e:
        print(f"Erro ao desativar a empresa: {e}")  # Para depuração
        flash(f"Erro ao desativar a empresa: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/ativar-empresa', methods=['POST'])
def ativar_empresa():
    empresa_inativa = request.form.get("empresa_inativa")  # Captura o nome da empresa inativa

    if not empresa_inativa:
        flash("Nenhuma empresa selecionada para ativar.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        # Buscar o id da empresa selecionada
        cursor = conn.cursor()
        cursor.execute("SELECT id_Empresa FROM Empresa WHERE Empresas = ?", (empresa_inativa,))
        id_empresa = cursor.fetchone()[0]

        # Atualizar o status da empresa na tabela Site_Empresa para ativa
        cursor.execute("UPDATE Site_Empresa SET Ativo = True WHERE id_Empresas = ?", (id_empresa,))
        conn.commit()

        flash(f"Empresa '{empresa_inativa}' ativada com sucesso!", "success")
    except Exception as e:
        print(f"Erro ao ativar a empresa: {e}")  # Para depuração
        flash(f"Erro ao ativar a empresa: {str(e)}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/programa-ferias', methods=['POST'])
def programa_ferias():
    nome = request.form.get('nome_ativo')
    data_inicio = request.form.get('data_inicio')
    data_fim = request.form.get('data_fim')
    siteempresa_id = request.form.get('siteempresa_id')  # Recebe o input hidden do site/empresa

    # Valida se os campos foram preenchidos
    if not nome or not data_inicio or not data_fim:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        # Converte as datas de início e fim para o formato datetime
        data_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
        data_fim = datetime.strptime(data_fim, '%Y-%m-%d')

        # Garante que a data de início não seja maior que a data de fim
        if data_inicio > data_fim:
            flash("A data de início não pode ser maior que a data de fim.", "error")
            return redirect(url_for('adiciona_presenca'))

        # Busca o ID do nome associado ao site e empresa
        cursor = conn.cursor()
        cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
        id_nome_result = cursor.fetchone()

        if id_nome_result is None:
            flash(f"Nome '{nome}' não encontrado para o site/empresa selecionado.", "error")
            return redirect(url_for('adiciona_presenca'))

        id_nome = id_nome_result[0]

        # Busca o ID da presença "FÉRIAS"
        cursor.execute("SELECT id_Presenca FROM Presenca WHERE Presenca = 'FÉRIAS'")
        id_presenca_result = cursor.fetchone()

        if id_presenca_result is None:
            flash("Tipo de presença 'FÉRIAS' não encontrado.", "error")
            return redirect(url_for('adiciona_presenca'))

        id_presenca = id_presenca_result[0]

        # Verifica se o total de dias de férias excede 30 dias
        cursor.execute("""
            SELECT COUNT(*) FROM Controle 
            WHERE id_Nome = ? AND id_Presenca = ? AND id_SiteEmpresa = ?
        """, (id_nome, id_presenca, siteempresa_id))
        total_dias_ferias = cursor.fetchone()[0]

        # Calcula o total de dias que o usuário quer adicionar
        dias_programados = (data_fim - data_inicio).days + 1

        if total_dias_ferias + dias_programados > 30:
            flash(f"O nome '{nome}' já tem {total_dias_ferias} dias de férias programados. "
                  f"Com esses novos {dias_programados} dias, o total excede o limite de 30 dias.", "error")
            return redirect(url_for('adiciona_presenca'))

        # Itera sobre cada dia no intervalo de datas
        current_date = data_inicio
        while current_date <= data_fim:
            # Não ignorar sábados e domingos para o tipo "FÉRIAS"
            cursor.execute("""
                INSERT INTO Controle (id_Nome, id_Presenca, Data, id_SiteEmpresa)
                VALUES (?, ?, ?, ?)
            """, (id_nome, id_presenca, current_date, siteempresa_id))
            current_date += timedelta(days=1)

        conn.commit()
        flash(f"Férias programadas com sucesso para {nome} de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}", "success")

    except Exception as e:
        flash(f"Erro ao programar férias: {e}", "error")

    return redirect(url_for('adiciona_presenca'))

@app.route('/desprogramar-ferias', methods=['POST'])
def desprogramar_ferias():
    nome = request.form.get('nome_ativo')
    data_inicio = request.form.get('data_inicio')
    data_fim = request.form.get('data_fim')
    siteempresa_id = request.form.get('siteempresa_id')  # Recebe o input hidden do site/empresa

    # Valida se os campos foram preenchidos
    if not nome or not data_inicio or not data_fim:
        flash("Por favor, preencha todos os campos.", "error")
        return redirect(url_for('adiciona_presenca'))

    try:
        # Converte as datas de início e fim para o formato datetime
        data_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
        data_fim = datetime.strptime(data_fim, '%Y-%m-%d')

        # Garante que a data de início não seja maior que a data de fim
        if data_inicio > data_fim:
            flash("A data de início não pode ser maior que a data de fim.", "error")
            return redirect(url_for('adiciona_presenca'))

        # Busca o ID do nome associado ao site e empresa
        cursor = conn.cursor()
        cursor.execute("SELECT id_Nomes FROM Nome WHERE Nome = ? AND id_SiteEmpresa = ?", (nome, siteempresa_id))
        id_nome_result = cursor.fetchone()

        if id_nome_result is None:
            flash(f"Nome '{nome}' não encontrado para o site/empresa selecionado.", "error")
            return redirect(url_for('adiciona_presenca'))

        id_nome = id_nome_result[0]

        # Busca o ID da presença "FÉRIAS"
        cursor.execute("SELECT id_Presenca FROM Presenca WHERE Presenca = 'FÉRIAS'")
        id_presenca_result = cursor.fetchone()

        if id_presenca_result is None:
            flash("Tipo de presença 'FÉRIAS' não encontrado.", "error")
            return redirect(url_for('adiciona_presenca'))

        id_presenca = id_presenca_result[0]

        # Itera sobre cada dia no intervalo de datas e remove as presenças "FÉRIAS"
        current_date = data_inicio
        while current_date <= data_fim:
            cursor.execute("""
                DELETE FROM Controle 
                WHERE id_Nome = ? AND id_Presenca = ? AND Data = ? AND id_SiteEmpresa = ?
            """, (id_nome, id_presenca, current_date, siteempresa_id))
            current_date += timedelta(days=1)

        conn.commit()
        flash(f"Férias desprogramadas com sucesso para {nome} de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}", "success")

    except Exception as e:
        flash(f"Erro ao desprogramar férias: {e}", "error")

    return redirect(url_for('adiciona_presenca'))

if __name__ == "__main__":
    # print('Runing on http://127.0.0.1/5000')
    app.run(debug=True)