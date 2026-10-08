# 📄 Leitor XML de NF-e → Contas a Receber

Aplicação desenvolvida em **Python** para automatizar a leitura de arquivos XML de **NF-e (Nota Fiscal Eletrônica)** e organizar os dados em uma planilha Excel de **Contas a Receber**.

O objetivo é reduzir o trabalho manual de cadastrar notas fiscais e suas respectivas parcelas na planilha financeira.

---

## 🚀 Funcionalidades

* 📂 Leitura de arquivos XML de NF-e
* 🔎 Extração automática dos principais dados da nota
* 📊 Organização das informações em uma planilha Excel
* 📅 Ordenação das contas pelo **dia de vencimento**, do menor para o maior
* 👤 Registro do nome do cliente
* 🗓️ Registro da data de emissão da NF-e
* 🧾 Registro do número da nota fiscal
* 💰 Identificação das parcelas da nota
* ⚡ Tratamento de notas à vista

---

## 📋 Informações extraídas

Para cada NF-e, o sistema organiza informações como:

| Informação          | Descrição                                       |
| ------------------- | ----------------------------------------------- |
| **Vencimento**      | Data de vencimento da parcela                   |
| **Nome**            | Nome do cliente                                 |
| **Data de emissão** | Data em que a NF-e foi emitida                  |
| **Nota**            | Número da NF-e e, quando aplicável, sua parcela |

---

## 🧾 Identificação das parcelas

Quando uma nota fiscal possui parcelas, o número da nota é combinado com o número correspondente da parcela.

Por exemplo, para a NF-e **5123**:

### Nota parcelada

```text
5123/001
5123/002
5123/003
```

Cada parcela recebe seu próprio registro e respectiva data de vencimento.

### Nota à vista

Quando a nota não possui parcelas, o sistema mantém apenas o número original:

```text
5123
```

---

## 📅 Ordenação por vencimento

As contas são organizadas de acordo com o **dia do vencimento**, começando pelo menor dia do mês e seguindo até o maior.

Por exemplo:

```text
01/10
03/10
05/10
10/10
15/10
22/10
28/10
```

Isso facilita a visualização das contas que precisam ser recebidas ao longo do mês.

---

## 🔄 Fluxo de funcionamento

O processo pode ser resumido da seguinte forma:

```text
XML da NF-e
     │
     ▼
Leitura do XML
     │
     ▼
Extração dos dados
     │
     ├── Cliente
     ├── Data de emissão
     ├── Número da NF-e
     └── Parcelas / vencimentos
     │
     ▼
Organização das informações
     │
     ▼
Ordenação por vencimento
     │
     ▼
Planilha Excel
     │
     ▼
Contas a Receber
```

---

## 🛠️ Tecnologias utilizadas

* **Python**
* **XML / ElementTree** — leitura e processamento dos arquivos XML
* **openpyxl** — criação e manipulação das planilhas Excel
* **Tkinter** — interface gráfica para interação com o programa
* **PyInstaller** — geração do executável para Windows

---

## 📁 Estrutura do projeto

Uma estrutura básica do projeto pode ser organizada assim:

```text
LeitorXML/
│
├── main.py
├── requirements.txt
├── README.md
│
├── xml_icon.ico
│
└── dist/
    └── LeitorXML.exe
```

---

## ▶️ Executando o projeto

Instale as dependências:

```bash
pip install -r requirements.txt
```

Depois execute:

```bash
python main.py
```

---

## 📦 Gerando o executável

Para gerar uma versão executável para Windows sem abrir o terminal:

```bash
pyinstaller --onefile --windowed --icon=xml_icon.ico main.py
```

O executável será criado dentro da pasta:

```text
dist/
```

---

## 🎯 Objetivo do projeto

Este projeto foi desenvolvido com o objetivo de **automatizar uma tarefa financeira repetitiva**, transformando os dados presentes nos XMLs das NF-e em registros organizados na planilha de **Contas a Receber**.

A automação facilita principalmente o controle de:

* clientes;
* notas fiscais;
* parcelas;
* datas de emissão;
* vencimentos;
* contas a receber.

