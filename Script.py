import json
import os
import datetime
import xlwings as xw

json_path = 'data.json'
modelo_path = 'Lista de IO Base.xlsx'
saida_dir = 'gerados'

os.makedirs(saida_dir, exist_ok=True)

with open(json_path, 'r', encoding='utf-8') as f:
    dados = json.load(f)

data = datetime.date.today()

projeto = input("Escolha o projeto (opções: OIS001, OIS002): ").upper()
base = input("Número da base (ex: 06): ").zfill(2)
arquivo = input("Número do arquivo (ex: 07): ").zfill(2)
revisao = input("Revisão (ex: 1): ")
execucao = input("Execução (ex: Kauan Barbosa -> K.B.): ").upper()

try:
    info_base = dados[projeto][base]
    info_doc = info_base["documentos"][arquivo]
except KeyError:
    print("❌ Projeto, base ou arquivo não encontrados.")
    exit()

substituicoes = {
    "<BASE_TITULO>": str(info_base.get("nome_base_titulo", "")),
    "<BASE>": str(info_base.get("nome_base", "")),
    "<ESTADO>": str(info_base.get("estado", "")),
    "<OT/SS/CC>": str(info_base.get("OT/SS/CC", "")),
    "<CODIGO_DOCUMENTO>": str(info_doc.get("codigo", "")),
    "<REV>": str(revisao),
    "<PAINEL>": str(info_base.get("painel", "")),
    "<DATA>": str(data),
    "<EXECUCAO>": str(execucao)
}

app = xw.App(visible=False)
wb = app.books.open(modelo_path)

for sheet in wb.sheets:
    used_range = sheet.used_range
    values = used_range.value

    for i, row in enumerate(values):
        for j, cell in enumerate(row):
            if isinstance(cell, str):
                for marcador, valor in substituicoes.items():
                    if marcador in cell:
                        cell = cell.replace(marcador, valor)
                values[i][j] = cell

    used_range.value = values

nome_arquivo = f"{info_doc['codigo']}_{revisao}.xlsx"
caminho_saida = os.path.join(saida_dir, nome_arquivo)

wb.save(caminho_saida)
wb.close()
app.quit()

print(f"✅ Documento gerado com sucesso: {caminho_saida}")
