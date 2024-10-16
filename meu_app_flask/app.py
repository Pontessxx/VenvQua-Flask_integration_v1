from flask import Flask, render_template, request, jsonify
import pandas as pd
import pyodbc
import json
import warnings
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

@app.route('/check_data', methods=['POST'])
def check_data():
    site = request.form.get('site')
    empresa = request.form.get('empresa')

    query = """
    SELECT COUNT(*) 
    FROM Site_Empresa 
    WHERE id_Sites = (SELECT id_Site FROM Site WHERE Sites = ?) 
      AND id_Empresas = (SELECT id_Empresa FROM Empresa WHERE Empresas = ?) 
      AND Ativo = True
    """
    cursor = conn.cursor()
    cursor.execute(query, (site, empresa))
    result = cursor.fetchone()

    if result and result[0] > 0:
        return jsonify({'hasData': True})
    else:
        return jsonify({'hasData': False})

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
                    # Filtrar pelo mês presente na data, comparando apenas o mês
                    selected_meses_numeric = [meses_dict[mes] for mes in selected_meses]
                    df = df[df['Data'].dt.strftime('%m').isin(selected_meses_numeric)]

                # Formatar a data para exibição
                df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')

                # Agrupar por tipo de presença e contar a frequência
                df_presenca = df.groupby('Presenca').size().reset_index(name='counts')
                labels = df_presenca['Presenca'].tolist()  # Tipos de presença
                values = df_presenca['counts'].tolist()    # Contagens de cada presença

                # Dados formatados para o gráfico de pizza
                pie_chart_data = json.dumps({
                    'labels': labels,
                    'values': values
                })
                
        except Exception as e:
            print(f"Erro ao consultar ou criar DataFrame: {e}")

    return render_template(
        "index.html",
        sites=sites,
        empresas=[e[1] for e in empresas],
        nomes=pd.read_sql("SELECT DISTINCT Nome FROM Nome", conn)['Nome'].tolist(),
        meses=meses_dict.keys(),  # Envia os nomes dos meses em português
        presencas=pd.read_sql("SELECT DISTINCT Presenca FROM Presenca", conn)['Presenca'].tolist(),
        selected_site=selected_site,
        selected_empresa=selected_empresa,
        selected_nomes=selected_nomes,
        selected_meses=selected_meses,
        selected_presenca=selected_presenca,
        data=df,
        pie_chart_data=pie_chart_data
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
    app.run(debug=True)
