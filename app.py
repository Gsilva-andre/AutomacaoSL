import pandas as pd
import urllib.parse
import re

# Caminho do seu arquivo
arquivo = "clientes.xlsx"

df = pd.read_excel(arquivo)

html = """
<html>
<head>
<meta charset="UTF-8">
<title>Clientes</title>
</head>
<body>

<h2>Lista de Clientes</h2>
"""

for _, row in df.iterrows():
    nome = row["razao social"]
    email = row["email"]
    telefone = str(row["telefone"])
    endereco = row["endereco"]

    # 🔥 Limpar telefone (remove tudo que não é número)
    telefone_limpo = re.sub(r'\D', '', telefone)

    # Email padrão
    assunto = urllib.parse.quote("Contato Comercial")
    corpo = urllib.parse.quote(
        f"Olá {nome}, tudo bem?\n\nGostaria de entrar em contato."
    )

    link_email = f"mailto:{email}?subject={assunto}&body={corpo}"
    link_whats = f"https://wa.me/55{telefone_limpo}"
    link_maps = f"https://www.google.com/maps/search/{urllib.parse.quote(endereco)}"

    html += f"""
    <div style="margin-bottom:20px;">
        <b>{nome}</b><br>
        <button onclick="window.location.href='{link_email}'">Email</button>
        <button onclick="window.open('{link_whats}')">WhatsApp</button>
        <button onclick="window.open('{link_maps}')">Maps</button>
        <hr>
    </div>
    """

html += "</body></html>"

with open("clientes.html", "w", encoding="utf-8") as f:
    f.write(html)

print("HTML gerado com sucesso!")