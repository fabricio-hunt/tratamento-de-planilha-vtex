# 📘 Pipeline de Tratamento e Divisão de Planilhas

Este projeto contém dois pipelines em Python desenvolvidos para **processar**, **sanitizar**, **gerar meta tags**, e **dividir grandes planilhas Excel** em arquivos menores.
Ele foi criado para lidar com planilhas extensas contendo HTML, descrições, atributos de e-commerce e outros dados que precisam ser preservados.

---

## 🚀 Funcionalidades Principais

### ✔ 1. Tratamento da Planilha e Geração de Meta Tags

Script baseado no arquivo enviado: `tratamento.txt`

* Geração automática de meta tags a partir do nome do produto.
* Criação da nova coluna **AE** preservando todo o conteúdo válido.
* Remoção **apenas** de caracteres ilegais do Excel.
* Preservação total de HTML, SKUs, códigos internos e demais dados.
* Salvamento seguro do arquivo, com detecção caso o arquivo já esteja aberto.
* Relatórios detalhados sobre limpeza e modificações.

---

### ✔ 2. Divisão da Planilha em Arquivos de até 4MB

Script baseado no arquivo enviado: `quebrar-branilhas.txt`

* Cálculo estimado do tamanho do Excel.
* Divisão automática em partes menores mantendo cabeçalho e estrutura.
* Ajuste dinâmico do número de linhas até caber no limite configurado.
* Criação de múltiplos arquivos organizados em diretório próprio.
* Relatórios completos com tamanho real de cada parte.

---

## 🧩 Tecnologias Utilizadas

* Python 3
* Pandas
* OpenPyXL
* XlsxWriter
* Regex
* Pathlib
* Math

---

## 📁 Estrutura Recomendada do Projeto

```
/pipeline-planilhas
│
├── src/
│   ├── tratamento.py
│   ├── dividir_planilha.py
│
├── input/
│   ├── dados-filtrados.xls
│   └── dados-filtrados_PROCESSADO.xlsx
│
├── output/
│   ├── dados-filtrados_PROCESSADO.xlsx
│   └── planilhas_divididas/
│
└── README.md
```

---

## 🛠 Como Executar

### Instale as dependências:

```bash
pip install pandas openpyxl xlsxwriter
```

### Execute o pipeline de tratamento:

```bash
python tratamento.py
```

### Execute o pipeline de divisão:

```bash
python dividir_planilha.py
```

> **Dica:** Para reuso, transforme os caminhos dos arquivos em parâmetros configuráveis.

---

## ⚙️ Configuração

Ambos os scripts utilizam caminhos fixos, como:

```
C:\Users\PC\Downloads\pipiline bemol farma\
```

Recomenda-se:

* Criar um arquivo `config.json`
* Ou permitir entrada via CLI (ex.: `--input arquivo.xlsx`)

Posso gerar isso automaticamente se desejar.

---

## 📌 Melhorias Futuras Sugeridas

* Adicionar interface de linha de comando com `argparse`.
* Criar logs persistentes (arquivo `.log`).
* Criar interface web local (Flask ou Streamlit).
* Criar testes automatizados com `pytest`.
* Criar versão executável `.exe` para Windows.

---

---

## 📄 Licença

Recomendação padrão:

```
MIT License
```
