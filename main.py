import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta

INPUT_EXCEL = "portfolio.xlsx"
OUTPUT_EXCEL = f"portfolio_atualizado-{datetime.now().date().strftime('%Y-%m-%d')}.xlsx"

def carregar_portfolio(caminho=INPUT_EXCEL):
    df = pd.read_excel(caminho)

    cols_esperadas = {"Ativo", "Quantidade", "Preço_Médio"}
    faltantes = cols_esperadas - set(df.columns)
    if faltantes:
        raise ValueError(f"Colunas faltando no Excel: {faltantes}. Esperado: {cols_esperadas}")

    df["Ativo"] = df["Ativo"].astype(str).str.strip()
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0)
    df["Preço_Médio"] = pd.to_numeric(df["Preço_Médio"], errors="coerce").fillna(0.0)

    if df["Ativo"].duplicated().any():
        df = (
            df.groupby("Ativo", as_index=False)
              .apply(lambda g: pd.Series({
                  "Quantidade": g["Quantidade"].sum(),
                  "Preço_Médio": np.average(g["Preço_Médio"], weights=g["Quantidade"])
              }))
              .reset_index()
              .drop(columns=["level_0"])
        )

    df = df[(df["Ativo"].str.len() > 0) & (df["Quantidade"] > 0)]
    return df

def baixar_precos(tickers):

    tickers = sorted(set(tickers))
    hist = yf.download(
        tickers=tickers,
        period="10d",
        interval="1d",
        group_by="ticker",
        auto_adjust=True,
        threads=True,
        progress=False
    )

    last_close = {}
    prev_close = {}

    # se um único ticker, o formato vem sem o nível por ticker
    if isinstance(hist.columns, pd.MultiIndex):

        for t in tickers:
            try:
                serie = hist[(t, "Close")].dropna()
                if len(serie) >= 2:
                    last_close[t] = float(serie.iloc[-1])
                    prev_close[t] = float(serie.iloc[-2])
                elif len(serie) == 1:
                    last_close[t] = float(serie.iloc[-1])
                    prev_close[t] = np.nan
                else:
                    last_close[t] = np.nan
                    prev_close[t] = np.nan
            except Exception:
                last_close[t] = np.nan
                prev_close[t] = np.nan
    else:
        serie = hist["Close"].dropna()
        t = tickers[0]
        if len(serie) >= 2:
            last_close[t] = float(serie.iloc[-1])
            prev_close[t] = float(serie.iloc[-2])
        elif len(serie) == 1:
            last_close[t] = float(serie.iloc[-1])
            prev_close[t] = np.nan
        else:
            last_close[t] = np.nan
            prev_close[t] = np.nan

    return last_close, prev_close

def coletar_dividendos_12m(tickers):

    out = {}
    cutoff = datetime.now() - timedelta(days=365)
    for t in set(tickers):
        try:
            div = yf.Ticker(t).dividends
            if div is None or len(div) == 0:
                out[t] = 0.0
            else:
                # filtra últimos 12 meses
                div_12m = div[div.index >= pd.to_datetime(cutoff)].sum()
                out[t] = float(div_12m) if pd.notna(div_12m) else 0.0
        except Exception:
            out[t] = 0.0
    return out

def calcular_metricas(df_base):
    tickers = df_base["Ativo"].tolist()

    last_close, prev_close = baixar_precos(tickers)
    div_12m = coletar_dividendos_12m(tickers)

    df = df_base.copy()
    df["Preço_Atual"]   = df["Ativo"].map(last_close)
    df["Preço_Anterior"] = df["Ativo"].map(prev_close)

    # Valores e P&L
    df["Valor_Investido"] = df["Quantidade"] * df["Preço_Médio"]
    df["Valor_Atual"]     = df["Quantidade"] * df["Preço_Atual"]
    df["P&L_R$"]          = df["Valor_Atual"] - df["Valor_Investido"]
    df["P&L_%"]           = np.where(df["Valor_Investido"] > 0,
                                     (df["Valor_Atual"] / df["Valor_Investido"] - 1) * 100,
                                     np.nan)

    # Variação diária
    df["Variação_Diária_%"] = np.where(
        df["Preço_Anterior"] > 0,
        (df["Preço_Atual"] / df["Preço_Anterior"] - 1) * 100,
        np.nan
    )
    df["P&L_Diário_R$"] = df["Quantidade"] * (df["Preço_Atual"] - df["Preço_Anterior"])

    # Dividendos últimos 12m (por ação * quantidade)
    df["Dividendos_12m_R$_por_ação"] = df["Ativo"].map(div_12m)
    df["Dividendos_12m_R$"] = df["Quantidade"] * df["Dividendos_12m_R$_por_ação"]

    # Peso na carteira
    valor_total = df["Valor_Atual"].sum(skipna=True)
    df["Peso_%"] = np.where(
        valor_total > 0,
        (df["Valor_Atual"] / valor_total) * 100,
        np.nan
    )

    # Ordena pelos maiores P&L
    df = df.sort_values(by="P&L_R$", ascending=False, na_position="last").reset_index(drop=True)

    # Resumo da carteira
    resumo = pd.DataFrame({
        "Métrica": [
            "Valor Investido (R$)",
            "Valor Atual (R$)",
            "P&L Total (R$)",
            "P&L Total (%)",
            "P&L Diário (R$)",
            "Total de Ativos",
            "Dividendos 12m (R$)"
        ],
        "Valor": [
            df["Valor_Investido"].sum(skipna=True),
            valor_total,
            df["P&L_R$"].sum(skipna=True),
            (valor_total / df["Valor_Investido"].sum(skipna=True) - 1) * 100 if df["Valor_Investido"].sum(skipna=True) > 0 else np.nan,
            df["P&L_Diário_R$"].sum(skipna=True),
            df["Ativo"].nunique(),
            df["Dividendos_12m_R$"].sum(skipna=True)
        ]
    })

    # Alocação
    alocacao = df[["Ativo", "Valor_Atual", "Peso_%"]].sort_values("Peso_%", ascending=False, na_position="last")
    alocacao = alocacao.reset_index(drop=True)

    # Top movers (variação diária)
    movers = df[["Ativo", "Variação_Diária_%", "P&L_Diário_R$"]].sort_values("Variação_Diária_%", ascending=False, na_position="last")
    movers = movers.reset_index(drop=True)

    return df, resumo, alocacao, movers

def salvar_excel(df_posicoes, df_resumo, df_alocacao, df_movers, caminho=OUTPUT_EXCEL):
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df_posicoes.to_excel(writer, sheet_name="Posições", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
        df_alocacao.to_excel(writer, sheet_name="Alocação", index=False)
        df_movers.to_excel(writer, sheet_name="Top Movers", index=False)

        for sheet, cols_fmt in [
            ("Posições", ["Preço_Médio", "Preço_Anterior", "Preço_Atual", "Valor_Investido", "Valor_Atual", "P&L_R$", "P&L_%", "Variação_Diária_%", "P&L_Diário_R$", "Dividendos_12m_R$_por_ação", "Dividendos_12m_R$", "Peso_%"]),
            ("Resumo", ["Valor"]),
            ("Alocação", ["Valor_Atual", "Peso_%"]),
            ("Top Movers", ["Variação_Diária_%", "P&L_Diário_R$"])
        ]:
            try:
                ws = writer.book[sheet]
                df_tmp = {
                    "Posições": df_posicoes,
                    "Resumo": df_resumo,
                    "Alocação": df_alocacao,
                    "Top Movers": df_movers
                }[sheet]
                for col_name in cols_fmt:
                    if col_name in df_tmp.columns:
                        col_idx = df_tmp.columns.get_loc(col_name) + 1  
                        for row in range(2, len(df_tmp) + 2): 
                            cell = ws.cell(row=row, column=col_idx)
                            if isinstance(cell.value, (int, float)) and cell.value is not None:
                                # porcentagens vs valores
                                if col_name.endswith("%"):
                                    cell.number_format = "0.00"
                                else:
                                    cell.number_format = "#,##0.00"
            except Exception:
                # se algo der errado, segue sem formatação
                pass

def main():
    df_base = carregar_portfolio(INPUT_EXCEL)
    df_posicoes, df_resumo, df_alocacao, df_movers = calcular_metricas(df_base)
    salvar_excel(df_posicoes, df_resumo, df_alocacao, df_movers, OUTPUT_EXCEL)
    # exibe um pequeno resumo no console para verificacao
    print("\n=== RESUMO DA CARTEIRA ===")
    print(df_resumo.to_string(index=False))
    print("\nArquivo salvo em:", OUTPUT_EXCEL)

if __name__ == "__main__":
    main()
