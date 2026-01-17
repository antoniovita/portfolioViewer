# Investment Portfolio Monitor with Python, Pandas, and Yahoo Finance

## Requirements

* Have **Python 3** installed on your machine.  
  👉 [Download Python here](https://www.python.org/downloads/)

---

## Features

* Reads an **Excel** file containing your investments.
* Automatically fetches **up-to-date market prices** using `yfinance`.
* Calculates:
  * Invested value
  * Current value
  * Return (%)
* Exports the results to a new Excel file with all calculations.
* Portfolio visualization in an **interactive browser-based dashboard**.

---

## Excel File Structure

Use the example Excel file as a base and place your portfolio data in it.

| Asset    | Quantity | Average_Price |
| -------- | -------- | ------------- |
| PETR4.SA | 100      | 27.50         |
| VALE3.SA | 50       | 65.00         |
| AAPL     | 10       | 150.00        |

* **Asset** → ticker symbol (Brazilian stocks end with `.SA`)
* **Quantity** → number of shares
* **Average_Price** → average purchase price per asset

**Follow the example format exactly**

---

## Virtual Environment Setup

### 1. Create a virtual environment

#### Mac / Linux

```bash
python3 -m venv venv
source venv/bin/activate
