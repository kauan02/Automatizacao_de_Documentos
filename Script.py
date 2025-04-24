import json
import os
import shutil
from openpyxl import load_workbook

# Caminhos
json_path = 'data.json'
modelo_path = 'Lista de IO Base.xlsx'
saida_dir = 'gerados'

# Cria pasta de saída se não existir
os.makedirs(saida_dir, exist_ok=True)

# Carrega JSON
with open(json_path, 'r', encoding='utf-8') as f:
    dados = json.load(f)

# Entradas do usuário
projeto = "OIS001"
base = input("Número da base (ex: 06): ").zfill(2)
arquivo = input("Número do arquivo (ex: 07): ").zfill(2)
revisao = input("Revisão (ex: 1): ")
data = input("Data (ex: 24/04/2025): ")
execucao = input("Execução (ex: Kauan Barbosa -> K.B.): ").upper()

# Busca dados no JSON
try:
    info_base = dados[projeto][base]
    info_doc = info_base["documentos"][arquivo]
except KeyError:
    print("❌ Projeto, base ou arquivo não encontrados.")
    exit()

# Dicionário com os valores a substituir
substituicoes = {
    "BASE_TITULO": str(info_base.get("nome_base_titulo", "")),
    "BASE": str(info_base.get("nome_base", "")),
    "ESTADO": str(info_base.get("estado", "")),
    "OT/SS/CC": str(info_base.get("OT/SS/CC", "")),
    "CODIGO_DOCUMENTO": str(info_doc.get("codigo", "")),
    "REV": str(revisao),
    "PAINEL": str(info_base.get("painel", "")),
    "DATA": str(data),
    "EXECUCAO": str(execucao)
}

# Copia o modelo original com a imagem para o novo nome
nome_arquivo = f"{info_doc['codigo']}_{revisao}.xlsx"
caminho_saida = os.path.join(saida_dir, nome_arquivo)
shutil.copy2(modelo_path, caminho_saida)  # Copia o arquivo preservando todos os dados, incluindo a imagem

# Abre o modelo copiado
wb = load_workbook(caminho_saida)

# Substituição nos textos das células
for sheet in wb.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for chave, valor in substituicoes.items():
                    marcador = f"<{chave}>"
                    if marcador in cell.value:
                        cell.value = cell.value.replace(marcador, valor)

# Salva o arquivo final no mesmo caminho
wb.save(caminho_saida)

print(f"✅ Documento gerado com sucesso: {caminho_saida}")
