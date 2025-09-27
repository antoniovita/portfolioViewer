# Monitor de Portfólio de Investimentos com Python, Pandas e Yahoo Finance

## Requisitos

* Ter o **Python 3** instalado na sua máquina.
  👉 [Baixe aqui o Python](https://www.python.org/downloads/)

---

## Funcionalidades

* Leitura de um arquivo **Excel** contendo seus investimentos.
* Coleta automática de **cotações atualizadas** usando o `yfinance`.
* Cálculo de:

  * Valor investido
  * Valor atual
  * Rentabilidade (%)
* Exportação para um novo arquivo Excel com todos os cálculos.
* Visualização do portfólio em um **dashboard interativo via navegador**.

---

## Estrutura do Arquivo Excel

Use o excel de exemplo como base e coloque seu portfólio nele.

| Ativo    | Quantidade | Preço_Médio |
| -------- | ---------- | ----------- |
| PETR4.SA | 100        | 27.50       |
| VALE3.SA | 50         | 65.00       |
| AAPL     | 10         | 150.00      |

* **Ativo** → código do ticker (ações brasileiras terminam em `.SA`)
* **Quantidade** → número de ações cotadas
* **Preço_Médio** → preço médio pago por ativo

**Siga o padrão do exemplo**

---

## Instalação e Ambiente Virtual

### 1. Criar ambiente virtual

#### Mac / Linux

```bash
python3 -m venv venv
source venv/bin/activate
```

#### Windows (PowerShell)

```powershell
python -m venv venv
.\venv\Scripts\activate
```

---

### 2. Instalar dependências

Com a venv ativa, instale as bibliotecas necessárias:

```bash
pip install pandas yfinance openpyxl
```

---

## Como Usar

Execute o script principal:

```bash
python main.py ou clique em cima do arquivo
```

Ele irá:

1. Ler o arquivo `portfolio.xlsx`.
2. Consultar as cotações mais recentes no Yahoo Finance.
3. Calcular valores atualizados e rentabilidade.
4. Gerar o arquivo `portfolio_atualizado<data-hoje>.xlsx`.

---

## Visualizador de Portfólio

Após rodar o script Python, você pode visualizar seu portfólio de forma interativa no navegador:

1. Abra o arquivo `index.html` na pasta do projeto.
2. Clique no botão para **anexar o Excel atualizado** (`portfolio_atualizado-AAAA-MM-DD.xlsx`).
3. Explore os gráficos e métricas do seu portfólio.

---

## Atualização

Sempre que você executar o `main.py`, ele atualizará automaticamente os dados do seu Excel e recalculará as estatísticas.