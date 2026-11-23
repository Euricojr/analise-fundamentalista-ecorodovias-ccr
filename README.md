# Análise Fundamentalista: EcoRodovias (ECOR3) vs. CCR (CCRO3 / MOTV3)

Repositório acadêmico contendo o código de extração de dados contábeis da **CVM** e de mercado da **B3**, planilhas de cálculo com fórmulas ativas e materiais para o trabalho de Análise Fundamentalista.

---

## 📌 Conteúdo do Projeto

* 📊 **`Trabalho_Analise_Fundamentalista_EcoRodovias_vs_CCR.xlsx`**:
  * **Aba 1 (Dados Primários):** Insumos contábeis brutos do Balanço Patrimonial (BP), DRE (12M) e mercado (Ações, Cotação e Dívida Líquida) para 2024, 2025 e 2026 LTM.
  * **Aba 2 (Indicadores Financeiros):** Quadro comparativo amarrado com fórmulas dinâmicas do Excel para os 13 indicadores obrigatórios.
* 📈 **`grafico_evolucao_acoes_ecor3_vs_motv3.png`**:
  * Gráfico comparativo de cotações das ações dos últimos 3 anos (2024 a 2026) pronto para inserção no PowerPoint/Word.
* 🐍 **`extrator_dados_cvm_b3.py`**:
  * Script em Python para extrair os dados e calcular os indicadores de **qualquer empresa da B3 / CVM**.
* ⚙️ **`gerar_excel_analise.py`**:
  * Script automatizado que gera o arquivo Excel com formatação visual e fórmulas ativas.

---

## 🚀 Como Executar e Puxar Dados de Outras Empresas

Se você  quiser puxar os dados de outras empresas (ex: `VALE3`, `PETR4`, `WEGE3`, etc.), execute:

```bash
# Para rodar com as empresas padrão (EcoRodovias e CCR):
python extrator_dados_cvm_b3.py

# Para rodar com outras empresas:
python extrator_dados_cvm_b3.py --tickers VALE3 PETR4

# Para exportar os resultados calculados para CSV:
python extrator_dados_cvm_b3.py --tickers ECOR3 CCRO3 --exportar_csv
```

---

## 📐 Indicadores Calculados

1. **Liquidez:** Liquidez Corrente (`Ativo Circulante / Passivo Circulante`).
2. **Rentabilidade:**
   * Margem Bruta (`Lucro Bruto / Receita Líquida`);
   * Margem Líquida (`Lucro Líquido / Receita Líquida`);
   * Giro do Ativo (`Receita Líquida / Ativo Total`);
   * ROA (`Lucro Líquido / Ativo Total`);
   * ROE (`Lucro Líquido / Patrimônio Líquido`).
3. **Endividamento:** Dívida / Patrimônio Líquido (`Dívida Bruta / PL`).
4. **Múltiplos de Mercado:**
   * Preço / Lucro (`P/L`);
   * Preço / Valor Patrimonial (`P/VPA`);
   * Preço / EBIT (`P/EBIT`);
   * EV / EBIT (`Enterprise Value / EBIT`);
   * Taxa de Dividend Yield (`DY`).

---

## 📋 Regra de Metodologia para 2026 (LTM / Projeção)

Conforme o enunciado do trabalho:
* **Balanço Patrimonial:** Dados da posição em **30/06/2026** (2T2026).
* **DRE (12 meses):** Acumulado móvel dos últimos 12 meses:  
  $$\text{DRE 2026} = \text{1T2026} + \text{2T2026} + \text{3T2025} + \text{4T2025}$$
