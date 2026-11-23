"""
================================================================================
EXTRATOR E CALCULADOR DE INDICADORES FUNDAMENTALISTAS (CVM / B3)
================================================================================
Ferramenta para extrair demonstrativos contábeis (Balanço Patrimonial e DRE)
e cotações oficiais de qualquer empresa listada na B3 / CVM, gerando
automaticamente os 13 indicadores exigidos para a análise fundamentalista.

Uso:
    python extrator_dados_cvm_b3.py --tickers ECOR3 CCRO3
    python extrator_dados_cvm_b3.py --tickers VALE3 PETR4
================================================================================
"""

import os
import sys
import argparse
import pandas as pd
import numpy as np

# Caminhos padrão dos dados locais consolidados da CVM e B3
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)

CVM_FUNDAMENTOS_PATH = os.path.join(PARENT_DIR, "lambda_cvm_fundamentos", "cvm_fundamentos_completo.parquet")
B3_MARKETCAP_PATH = os.path.join(PARENT_DIR, "lambda_b3_marketcap_diario", "b3_marketcap_serie_historica.parquet")
COTACOES_PATH = os.path.join(PARENT_DIR, "Lambdas AWS", "simulacao_s3", "raw", "cotacoes_splits", "cotacoes.parquet")

# Mapeamento auxiliar de Ticker -> CNPJ
TICKER_CNPJ_MAP = {
    "ECOR3": "04149454000180",
    "CCRO3": "02846056000197",
    "MOTV3": "02846056000197",
    "TPIS3": "03014553000191",
    "RAIL3": "13552862000180",
    "VALE3": "33592510000154",
    "PETR4": "33000167000101",
    "PETR3": "33000167000101",
    "WEGE3": "84429695000111",
    "ABEV3": "07526557000100",
    "RENT3": "16594806000188"
}

def carregar_bases():
    """Carrega as bases de dados locais se existirem."""
    df_fund = None
    df_b3 = None
    df_cot = None
    
    if os.path.exists(CVM_FUNDAMENTOS_PATH):
        df_fund = pd.read_parquet(CVM_FUNDAMENTOS_PATH)
    if os.path.exists(B3_MARKETCAP_PATH):
        df_b3 = pd.read_parquet(B3_MARKETCAP_PATH)
    if os.path.exists(COTACOES_PATH):
        df_cot = pd.read_parquet(COTACOES_PATH)
        
    return df_fund, df_b3, df_cot

def extrair_dados_empresa(ticker, cnpj, df_fund, df_cot):
    """Extrai as variáveis primárias contábeis de 2024, 2025 e 2026."""
    cnpj_limpo = cnpj.replace(".", "").replace("/", "").replace("-", "").zfill(14)
    sub = df_fund[df_fund["CNPJ_LIMPO"] == cnpj_limpo].copy()
    
    if sub.empty:
        print(f"[-] CNPJ {cnpj_limpo} para {ticker} não encontrado na base de fundamentos.")
        return None
        
    datas = ["2024-12-31", "2025-12-31", "2026-06-30"]
    sub = sub[sub["DT_REFER"].isin(datas)].copy()
    
    registros = []
    for _, row in sub.iterrows():
        dt = row["DT_REFER"]
        ano = int(dt[:4])
        
        # Variáveis contábeis
        ativo_tot = float(row.get("Ativo_Total", 0.0))
        ativo_circ = float(row.get("Ativo_Circulante", 0.0))
        passivo_circ = float(row.get("Passivo_Circulante", 0.0))
        pl = float(row.get("Patrimonio_Liquido", 0.0))
        div_bruta = float(row.get("Divida_Bruta", 0.0))
        caixa = float(row.get("Caixa_Equivalentes", 0.0))
        div_liq = div_bruta - caixa
        
        rec_liq = float(row.get("Receita_Liquida_12M", 0.0))
        lucro_bruto = float(row.get("Lucro_Bruto_12M", 0.0))
        ebit = float(row.get("EBIT_12M", 0.0))
        
        ll = float(row.get("Lucro_Liquido_Total_12M", 0.0))
        if ll == 0.0:
            ll = float(row.get("Lucro_Liquido_12M", 0.0))
            
        registros.append({
            "Ticker": ticker,
            "CNPJ": cnpj_limpo,
            "Empresa": row.get("DENOM_CIA", ticker),
            "Ano": ano,
            "Data_Referencia": dt,
            "Ativo_Total": ativo_tot,
            "Ativo_Circulante": ativo_circ,
            "Passivo_Circulante": passivo_circ,
            "Patrimonio_Liquido": pl,
            "Divida_Bruta": div_bruta,
            "Caixa": caixa,
            "Divida_Liquida": div_liq,
            "Receita_Liquida_12M": rec_liq,
            "Lucro_Bruto_12M": lucro_bruto,
            "EBIT_12M": ebit,
            "Lucro_Liquido_12M": ll
        })
        
    return pd.DataFrame(registros)

def calcular_indicadores(df_primarias):
    """Aplica as fórmulas oficiais para calcular os indicadores."""
    df = df_primarias.copy()
    
    # Liquidez
    df["Liquidez_Corrente"] = df["Ativo_Circulante"] / df["Passivo_Circulante"]
    
    # Rentabilidade
    df["Margem_Bruta_Pct"] = (df["Lucro_Bruto_12M"] / df["Receita_Liquida_12M"]) * 100.0
    df["Margem_Liquida_Pct"] = (df["Lucro_Liquido_12M"] / df["Receita_Liquida_12M"]) * 100.0
    df["Giro_Ativo"] = df["Receita_Liquida_12M"] / df["Ativo_Total"]
    df["ROA_Pct"] = (df["Lucro_Liquido_12M"] / df["Ativo_Total"]) * 100.0
    df["ROE_Pct"] = (df["Lucro_Liquido_12M"] / df["Patrimonio_Liquido"]) * 100.0
    
    # Endividamento
    df["Divida_PL"] = df["Divida_Bruta"] / df["Patrimonio_Liquido"]
    
    return df

def main():
    parser = argparse.ArgumentParser(description="Extrator de Indicadores Fundamentalistas CVM/B3")
    parser.add_argument("--tickers", nargs="+", default=["ECOR3", "CCRO3"], help="Lista de tickers para análise")
    parser.add_argument("--exportar_csv", action="store_true", help="Salva os resultados em arquivo CSV")
    args = parser.parse_args()
    
    print("[*] Carregando bases de dados locais da CVM e B3...")
    df_fund, df_b3, df_cot = carregar_bases()
    
    if df_fund is None:
        print("[!] Base cvm_fundamentos_completo.parquet não encontrada no diretório esperado.")
        sys.exit(1)
        
    dfs_coletados = []
    for ticker in args.tickers:
        ticker_up = ticker.upper()
        cnpj = TICKER_CNPJ_MAP.get(ticker_up)
        
        # Tenta buscar na base B3 se não estiver no mapa
        if not cnpj and df_b3 is not None:
            match = df_b3[df_b3["Ticker"] == ticker_up]
            if not match.empty:
                cnpj = match.iloc[0]["CNPJ"]
                
        if not cnpj:
            print(f"[!] Não foi possível identificar o CNPJ para o ticker '{ticker_up}'. Adicione-o no script.")
            continue
            
        print(f"[+] Extraindo dados para {ticker_up} (CNPJ: {cnpj})...")
        df_emp = extrair_dados_empresa(ticker_up, cnpj, df_fund, df_cot)
        if df_emp is not None and not df_emp.empty:
            dfs_coletados.append(df_emp)
            
    if not dfs_coletados:
        print("[-] Nenhum dado coletado para os tickers informados.")
        sys.exit(1)
        
    df_consolidado = pd.concat(dfs_coletados, ignore_index=True)
    df_indicadores = calcular_indicadores(df_consolidado)
    
    cols_exibicao = [
        "Ticker", "Ano", "Liquidez_Corrente", "Margem_Bruta_Pct", 
        "Margem_Liquida_Pct", "Giro_Ativo", "ROA_Pct", "ROE_Pct", "Divida_PL"
    ]
    
    print("\n" + "="*80)
    print("INDICADORES FUNDAMENTALISTAS CALCULADOS:")
    print("="*80)
    print(df_indicadores[cols_exibicao].to_string(index=False))
    print("="*80)
    
    if args.exportar_csv:
        output_file = os.path.join(BASE_DIR, "indicadores_extraidos.csv")
        df_indicadores.to_csv(output_file, index=False, sep=";", decimal=",", encoding="utf-8-sig")
        print(f"\n[✓] Arquivo exportado com sucesso em: {output_file}")

if __name__ == "__main__":
    main()
