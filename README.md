# 📄 Gerador de Documentos Automático

Projeto desenvolvido para facilitar a edição de documentos do Word (.docx) e Excel (.xlsx) utilizando dados armazenados em um banco de dados JSON.

A ferramenta permite buscar informações de projetos, bases e arquivos e gerar documentos personalizados com base em modelos previamente definidos.

## ✨ Funcionalidades

- Buscar dados de projetos no banco JSON
- Substituir variáveis em documentos Word e Excel
- Gerar documentos finais automaticamente
- Organização em estrutura de pastas por projeto e base
- Expansível para novos tipos de documentos e dados

## 🚀 Tecnologias Utilizadas

- [Python](https://www.python.org/) - Linguagem principal
- [python-docx](https://python-docx.readthedocs.io/en/latest/) - Manipulação de arquivos .docx
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) - Manipulação de arquivos .xlsx
- [JSON](https://www.json.org/json-pt.html) - Armazenamento de dados

## 📚 Pré-requisitos

Antes de começar, você vai precisar ter instalado em sua máquina:

- [Python 3.10+](https://www.python.org/downloads/)
- pip (gerenciador de pacotes Python)

Além disso, será necessário instalar as dependências do projeto:

```bash
pip install -r requirements.txt
```
## 🛠️ Como rodar o projeto

Clone o repositório:
```bash
git clone https://github.com/kauan02/Automatizacao_de_Documentos
```

Acesse a pasta do projeto:
```bash
cd Automatizacao_de_Documentos
```

Execute o script principal:
```bash
script.py
```
O sistema irá solicitar:

- ID do Projeto
- ID da Base
- ID do Arquivo

Após informar os dados, os documentos serão gerados automaticamente nas pastas corretas.

# 📁 Estrutura de Pastas
```bash
Automatizacao_de_Documentos/
├── Documentos_Base/
├── gerados/
├── data.json
├── script.py
└── README-PTBR.md
```
- data.json → Banco de dados com informações dos projetos.

- script.py → Script principal do sistema.

# 💡 Como contribuir
- Faça um fork do projeto.

- Crie uma nova branch com a sua feature (git checkout -b minha-feature).

- Faça o commit das suas alterações (git commit -m 'feat: Minha nova feature').

- Faça o push para a sua branch (git push origin minha-feature).

- Abra um Pull Request.


Feito com ❤️ por Kauan Barbosa Rezende 👨‍💻
