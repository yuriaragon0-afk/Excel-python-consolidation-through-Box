# 📊 Automação de Consolidação de Relatórios Excel (Python + xlwings)

Este projeto automatiza o processo de **consolidação de múltiplos relatórios Excel**, extraindo informações de abas específicas, unificando os dados em um arquivo de consolidação e executando **macros VBA** para gerar um relatório final formatado.

O projeto foi inspirado em uma automação real usada em ambiente financeiro corporativo, mas **todos os caminhos e dados foram substituídos por exemplos genéricos e fictícios** para preservar a confidencialidade da empresa.

---

## 🚀 Funcionalidades

- 🔄 Consolidação automática de vários arquivos Excel  
- 📑 Leitura de abas e intervalos específicos  
- 📦 Unificação de dados em um arquivo mestre  
- ⚙️ Execução de macros Excel diretamente via Python  
- 📈 Geração de um relatório final consolidado  

---

## 🧰 Tecnologias Utilizadas

- **Python 3.x**
- pandas — manipulação de dados  
- openpyxl — leitura e gravação de arquivos Excel  
- xlwings — automação do Excel e execução de macros  

---

## ⚙️ Estrutura do Projeto

```
seu-projeto/
├── BPT bridges dummy.ipynb      # Notebook principal (código da automação)
├── requirements.txt             # Dependências do projeto
├── .gitignore                   # Itens ignorados pelo Git
└── README.md                    # Este arquivo :)
```

---

## 🧩 Como Executar o Projeto

### 1️⃣ Clonar o repositório

```bash
git clone https://github.com/yuriaragon0-afk/Excel-python-consolidation-through-Box
cd Excel-python-consolidation-through-Box
```

### 2️⃣ Criar ambiente virtual (opcional)

```bash
python -m venv venv
venv\Scripts\activate   # Windows
# ou
source venv/bin/activate   # macOS/Linux
```

### 3️⃣ Instalar dependências

```bash
pip install -r requirements.txt
```

### 4️⃣ Configurar os caminhos no código

No início do notebook (ou do script Python), edite os caminhos conforme sua estrutura local:

```python
folder_path = r"C:/Exemplo/Relatorios/"
consolidation_path = r"C:/Exemplo/Consolidado/consolidado.xlsx"
source_sheet_name = "Resumo"
macro_name = "ExecutarConsolidacao"
```

Esses caminhos apontam para onde estão os arquivos de entrada e onde o consolidado será salvo.

---

## ▶️ Execução

Se estiver usando o **notebook**:
1. Abra `BPT bridges dummy.ipynb` no Jupyter ou VSCode  
2. Execute as células em sequência  

Se quiser transformar em **script Python**:
```bash
python consolidacao.py
```

O script:
1. Lê todos os arquivos Excel da pasta indicada  
2. Copia os dados das abas especificadas  
3. Consolida tudo em um único arquivo  
4. Executa a macro indicada  
5. Gera o relatório final consolidado  

---

## 📦 Exemplo de Estrutura de Dados

```
data/
├── relatorio_analista1.xlsx
├── relatorio_analista2.xlsx
└── relatorio_analista3.xlsx
```

---

## 💡 Observação: uso opcional de `.env`

Se quiser deixar o código mais flexível e seguro (boa prática profissional),  
você pode armazenar os caminhos e nomes de abas em um arquivo `.env` e ler com a biblioteca `python-dotenv`:

```python
from dotenv import load_dotenv
import os

load_dotenv()

folder_path = os.getenv("FOLDER_PATH")
consolidation_path = os.getenv("CONSOLIDATION_PATH")
source_sheet_name = os.getenv("SOURCE_SHEET_NAME")
macro_name = os.getenv("MACRO_NAME")
```

Exemplo de `.env`:
```
FOLDER_PATH=C:/Exemplo/Relatorios/
CONSOLIDATION_PATH=C:/Exemplo/Consolidado/consolidado.xlsx
SOURCE_SHEET_NAME=Resumo
MACRO_NAME=ExecutarConsolidacao
```

Mas o uso é **opcional** — o código também funciona com os caminhos definidos diretamente no script.

---

## ⚠️ Aviso de Confidencialidade

Este projeto foi inspirado em uma automação corporativa real, porém **todos os dados, nomes e caminhos foram substituídos por exemplos genéricos**.  
Nenhum conteúdo sensível ou confidencial está incluído neste repositório.

---

## 👤 Autor

**[Yuri Aragon]**  
Analista Financeiro | Python | Excel | Automação de Processos  
📧 [yuriaragon0@gmail.com] 
🌐 [https://www.linkedin.com/in/yuriaragon/]

---

## 🏷️ Licença

Distribuído sob a licença MIT. Consulte o arquivo `LICENSE` (opcional) para mais informações.
