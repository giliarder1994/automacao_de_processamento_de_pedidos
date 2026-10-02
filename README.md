# 🤖 Automação de Processamento de Pedidos em PDF

Projeto desenvolvido em **Python** com o objetivo de automatizar uma tarefa operacional realizada no ambiente de trabalho: **extrair informações de pedidos a partir de arquivos PDF e organizar os dados automaticamente em uma planilha Excel**.

Este projeto faz parte de uma automação maior que está sendo desenvolvida por etapas. A ideia final é reduzir tarefas manuais, diminuir erros de digitação e agilizar o processo de tratamento das informações dos pedidos.

---

## 🎯 Objetivo

Atualmente, o processo automatizado realiza:

**PDF → Extração das informações → Excel**

O script identifica automaticamente informações relevantes dentro dos documentos PDF e registra os dados em uma planilha Excel.

### Informações extraídas

* 🏢 Nome da empresa
* 💰 Valor total dos produtos
* 💵 Valor total da nota
* 📄 Nome do arquivo PDF processado

---

## 🚀 Próximas etapas

Este projeto está sendo desenvolvido de forma incremental.

A primeira etapa foi criar a automação responsável pela leitura dos PDFs e organização das informações em Excel.

### Etapa atual

* [x] Ler arquivos PDF
* [x] Extrair informações utilizando expressões regulares
* [x] Identificar empresa
* [x] Extrair valor total dos produtos
* [x] Extrair valor total da nota
* [x] Criar planilha Excel automaticamente
* [x] Adicionar múltiplos PDFs na mesma planilha
* [x] Processar uma pasta inteira de PDFs
* [ ] Automatizar geração do e-mail
* [ ] Integrar com Gmail
* [ ] Utilizar os dados extraídos para montar o corpo do e-mail
* [ ] Automatizar o envio do e-mail
* [ ] Melhorar tratamento de erros e validações

A ideia é chegar futuramente ao seguinte fluxo:

```text
              ┌──────────────┐
              │   PDF/PDFs   │
              └──────┬───────┘
                     │
                     ▼
          ┌─────────────────────┐
          │  Leitura do PDF     │
          │     pdfplumber       │
          └──────────┬──────────┘
                     │
                     ▼
          ┌─────────────────────┐
          │ Extração dos dados  │
          │      Regex          │
          └──────────┬──────────┘
                     │
                     ▼
          ┌─────────────────────┐
          │      Excel          │
          │    openpyxl         │
          └──────────┬──────────┘
                     │
                     ▼
          ┌─────────────────────┐
          │ Geração do e-mail   │
          └──────────┬──────────┘
                     │
                     ▼
          ┌─────────────────────┐
          │       Gmail         │
          └─────────────────────┘
```

---

## 🛠️ Tecnologias utilizadas

* **Python**
* **pdfplumber** — leitura e extração de texto dos arquivos PDF
* **openpyxl** — criação e atualização da planilha Excel
* **re (Regular Expressions)** — identificação das informações dentro do texto
* **pathlib** — manipulação de arquivos e diretórios
* **os** — verificação e manipulação de arquivos
* **sys** — configuração da saída do terminal

---

## 📂 Estrutura atual

```text
automacao-pdf-python/
│
├── pdf_para_excel.py
├── resultado_pedidos.xlsx
├── README.md
└── pdfs/
    ├── pedido_01.pdf
    ├── pedido_02.pdf
    └── pedido_03.pdf
```

> A estrutura pode ser reorganizada conforme novas etapas da automação forem implementadas.

---

## ⚙️ Como funciona

### 1. Localização dos PDFs

O programa pode receber:

* Um único arquivo PDF
* Uma pasta contendo vários arquivos PDF

Exemplo:

```python
processar_pdfs(
    pasta_ou_arquivo=r"C:\Users\Usuario\Desktop\automacao_pdf_python",
    caminho_excel="resultado_pedidos.xlsx"
)
```

---

### 2. Leitura do PDF

A biblioteca `pdfplumber` é utilizada para abrir o documento e extrair seu conteúdo textual.

O script analisa as páginas do PDF começando pela última página e seguindo em direção às primeiras.

```python
with pdfplumber.open(caminho_pdf) as pdf:
    total_paginas = len(pdf.pages)

    paginas_para_analisar = list(reversed(range(total_paginas)))
```

Essa abordagem foi utilizada porque as informações necessárias para o processo podem estar localizadas nas páginas finais do documento.

---

### 3. Extração das informações

Depois que o texto é extraído, são utilizadas **expressões regulares (Regex)** para localizar os dados necessários.

Por exemplo, para encontrar o valor total dos produtos:

```python
r'Valor\s+Total\s+dos\s+Produtos\s*=>\s*([\d.,]+)'
```

E para encontrar o valor total da nota:

```python
r'Valor\s+Total\s+da\s+Nota\s*=>\s*([\d.,]+)'
```

---

## 📊 Geração do Excel

Após a extração, os dados são armazenados em uma planilha utilizando `openpyxl`.

A planilha possui atualmente as seguintes colunas:

| Arquivo PDF   | Nome da Empresa   | Valor Total dos Produtos | Valor Total da Nota |
| ------------- | ----------------- | ------------------------ | ------------------- |
| pedido_01.pdf | Empresa Exemplo   | 1.500,00                 | 1.620,00            |
| pedido_02.pdf | Empresa Exemplo 2 | 850,00                   | 920,00              |

Caso a planilha já exista, o programa **não cria uma nova planilha**. Ele carrega o arquivo existente e adiciona uma nova linha com os dados do PDF processado.

---

## 📁 Processamento de vários PDFs

Uma das funcionalidades implementadas é a possibilidade de informar uma pasta inteira.

```python
elif caminho.is_dir():
    lista_pdfs = list(caminho.glob("*.pdf"))
```

Dessa forma, vários pedidos podem ser processados automaticamente em sequência.

Exemplo:

```text
📄 pedido_001.pdf
📄 pedido_002.pdf
📄 pedido_003.pdf
📄 pedido_004.pdf
```

O resultado será consolidado na mesma planilha:

```text
resultado_pedidos.xlsx
```

---

## 🖥️ Exemplo de execução

Ao executar o programa, o terminal apresenta o andamento do processamento:

```text
🔍 Encontrados 3 PDF(s) para processar...

📄 Processando: pedido_001.pdf
   → Empresa: EMPRESA EXEMPLO
   → Valor Produtos: 1.500,00
   → Valor Nota: 1.620,00

📄 Processando: pedido_002.pdf
   → Empresa: EMPRESA EXEMPLO 2
   → Valor Produtos: 850,00
   → Valor Nota: 920,00

🎉 Processamento finalizado!
```

---

## 🔎 Tratamento de informações não encontradas

Caso alguma informação não seja localizada no PDF, o programa retorna:

```text
NÃO ENCONTRADO
```

Isso permite identificar posteriormente quais documentos precisam de uma análise manual ou de uma melhoria na regra de extração.

---

## 💡 Por que este projeto foi desenvolvido?

A ideia surgiu a partir de uma tarefa operacional real que envolve:

1. Receber documentos PDF;
2. Abrir os arquivos individualmente;
3. Procurar informações específicas;
4. Copiar os dados;
5. Organizar as informações;
6. Utilizar esses dados posteriormente para comunicação por e-mail.

Em vez de realizar todas essas etapas manualmente, o projeto busca transformar o processo em um fluxo automatizado.

Além de resolver um problema prático, o projeto também serve como aplicação dos conhecimentos de **Python, manipulação de arquivos, processamento de dados, Regex e automação de tarefas**.

---

## 📚 Conceitos praticados

Durante o desenvolvimento foram utilizados conceitos como:

* Manipulação de arquivos
* Leitura de PDFs
* Extração e tratamento de texto
* Expressões regulares (Regex)
* Estruturas condicionais
* Funções
* Dicionários
* Laços de repetição
* Tratamento de exceções
* Manipulação de diretórios
* Criação e edição de arquivos Excel
* Organização de dados
* Automação de tarefas

---

## 🔮 Visão futura

O objetivo final é transformar o projeto em uma automação completa:

```text
📄 PDF
   ↓
🔎 Extração automática
   ↓
📊 Organização dos dados
   ↓
✉️ Geração automática do e-mail
   ↓
📧 Gmail
   ↓
✅ Processo concluído
```

Com isso, uma tarefa que atualmente exige várias etapas manuais poderá ser executada de forma muito mais rápida e padronizada.

---

## 👨‍💻 Autor

**Giliarde da Silva Rodrigues**

Estudante de Engenharia de Software com foco em desenvolvimento e automação utilizando Python.

[GitHub](https://github.com/)
