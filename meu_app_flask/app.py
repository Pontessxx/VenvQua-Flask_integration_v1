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

# Dicionário para os meses
meses_dict = {
    "01": "Janeiro",
    "02": "Fevereiro",
    "03": "Março",
    "04": "Abril",
    "05": "Maio",
    "06": "Junho",
    "07": "Julho",
    "08": "Agosto",
    "09": "Setembro",
    "10": "Outubro",
    "11": "Novembro",
    "12": "Dezembro"
}

@app.route("/", methods=["GET", "POST"])
def index():
    query_sites = "SELECT DISTINCT Sites FROM Site"
    sites = pd.read_sql(query_sites, conn)['Sites'].tolist()

    query_nomes = "SELECT DISTINCT Nome FROM Nome"
    nomes = pd.read_sql(query_nomes, conn)['Nome'].tolist()

    # Obter meses distintos como números
    query_meses = "SELECT DISTINCT FORMAT(Data, 'MM/yyyy') AS Mes FROM Controle"
    meses_numeros = pd.read_sql(query_meses, conn)['Mes'].tolist()

    # Mapear os meses para os nomes
    meses = {mes: meses_dict[mes.split('/')[0]] for mes in meses_numeros}

    selected_site = request.form.get("site")
    selected_nomes = request.form.getlist("nomes")
    selected_meses = request.form.getlist("meses")  # Captura os meses selecionados
    empresas = []
    empresa_opcoes = []
    df = None

    if selected_site:
        # Após selecionar o site, carregar empresas
        site_id = get_site_id(selected_site)
        empresas = get_empresas(site_id)
        empresa_opcoes = [empresa[1] for empresa in empresas]

        selected_empresa = request.form.get("empresa")
        
        if selected_empresa:
            empresa_id = get_empresa_id(selected_empresa, empresas)
            siteempresa_id = get_siteempresa_id(site_id, empresa_id)

            # Consulta inicial com o site e empresa selecionados
            query = f"""
            SELECT Nome.Nome, Presenca.Presenca, 
                   FORMAT(Controle.Data, 'dd/mm/yyyy') AS Data
            FROM Presenca 
            INNER JOIN (Nome 
            INNER JOIN Controle ON Nome.id_Nomes = Controle.id_Nome) 
            ON Presenca.id_Presenca = Controle.id_Presenca
            WHERE Controle.id_SiteEmpresa = ?
            """
            params = [siteempresa_id]

            # Se nomes forem selecionados, adicionar filtro extra
            if selected_nomes:
                query += f" AND Nome.Nome IN ({','.join(['?'] * len(selected_nomes))})"
                params += selected_nomes

            # Se meses forem selecionados, adicionar filtro extra
            if selected_meses:
                query += f" AND FORMAT(Controle.Data, 'MM/yyyy') IN ({','.join(['?'] * len(selected_meses))})"
                params += selected_meses

            df = pd.read_sql(query, conn, params=params)
        else:
            # Exibir a tabela sem filtro de nomes (apenas site e empresa)
            query = f"""
            SELECT Nome.Nome, Presenca.Presenca, 
                   FORMAT(Controle.Data, 'dd/mm/yyyy') AS Data
            FROM Presenca 
            INNER JOIN (Nome 
            INNER JOIN Controle ON Nome.id_Nomes = Controle.id_Nome) 
            ON Presenca.id_Presenca = Controle.id_Presenca
            WHERE Controle.id_SiteEmpresa = ?
            """
            params = [siteempresa_id]
            df = pd.read_sql(query, conn, params=params)

    return render_template(
        "index.html", sites=sites, empresas=empresa_opcoes, nomes=nomes,
        meses=meses.items(), selected_site=selected_site, selected_empresa=selected_empresa, 
        selected_nomes=selected_nomes, selected_meses=selected_meses, data=df
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
