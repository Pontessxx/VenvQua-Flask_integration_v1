from flask import Flask, render_template, request
import pandas as pd
import pyodbc

app = Flask(__name__)

# Configuração da conexão com o banco de dados Access
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\\Users\\Henrique\\Downloads\\Controle.accdb;"
)
conn = pyodbc.connect(conn_str)

# Dicionário para os meses em português
meses_dict = {
    "January": "Janeiro",
    "February": "Fevereiro",
    "March": "Março",
    "April": "Abril",
    "May": "Maio",
    "June": "Junho",
    "July": "Julho",
    "August": "Agosto",
    "September": "Setembro",
    "October": "Outubro",
    "November": "Novembro",
    "December": "Dezembro"
}

@app.route("/", methods=["GET", "POST"])
def index():
    query_sites = "SELECT DISTINCT Sites FROM Site"
    sites = pd.read_sql(query_sites, conn)['Sites'].tolist()

    selected_site = request.form.get("site")
    selected_empresa = request.form.get("empresa")
    selected_nomes = request.form.getlist("nomes")
    selected_meses = request.form.getlist("meses")
    selected_presenca = request.form.getlist("presenca")

    empresas = []
    if selected_site:
        empresas = get_empresas(get_site_id(selected_site))

    # Consulta inicial da tabela com filtros
    query = """
    SELECT Nome.Nome, Presenca.Presenca, Controle.Data
    FROM Presenca 
    INNER JOIN (Nome 
    INNER JOIN Controle ON Nome.id_Nomes = Controle.id_Nome) 
    ON Presenca.id_Presenca = Controle.id_Presenca
    """
    df = pd.read_sql(query, conn)

    if selected_nomes:
        df = df[df['Nome'].isin(selected_nomes)]
    if selected_meses:
        df = df[df['Mês'].isin(selected_meses)]
    if selected_presenca:
        df = df[df['Presenca'].isin(selected_presenca)]

    df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')

    return render_template(
        "index.html", 
        sites=sites, 
        empresas=[e[1] for e in empresas], 
        nomes=pd.read_sql("SELECT DISTINCT Nome FROM Nome", conn)['Nome'].tolist(), 
        meses=meses_dict.values(), 
        presencas=pd.read_sql("SELECT DISTINCT Presenca FROM Presenca", conn)['Presenca'].tolist(),
        selected_site=selected_site, 
        selected_empresa=selected_empresa,
        selected_nomes=selected_nomes, 
        selected_meses=selected_meses, 
        selected_presenca=selected_presenca,
        data=df
    )


# Funções de ajuda não alteradas
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
    cursor = conn.cursor()
    query = """
    SELECT id_SiteEmpresa FROM Site_Empresa 
    WHERE id_Sites = ? AND id_Empresas = ? AND Ativo = True
    """
    cursor.execute(query, (site_id, empresa_id))
    result = cursor.fetchone()
    return result[0] if result else None

if __name__ == "__main__":
    app.run(debug=True)
