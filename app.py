import streamlit as st
import pandas as pd
import re
from io import BytesIO  # Para criar o arquivo Excel em memória
import sys

# =================================================================================
# FUNÇÃO DE LÓGICA (Adaptada para Streamlit)
# =================================================================================
# Esta é a sua função original, mas adaptada para ler "file objects"
# em vez de "file paths" (caminhos de arquivo).
def processar_dados_streamlit(orcamento_file, gastos_file, ano, periodo_inicial, periodo_final):
    """
    Função adaptada para o Streamlit que:
    1. Lê "file objects" (de upload) em vez de caminhos.
    2. Limpa colunas "Unnamed".
    3. Substitui os valores do orçamento pelos gastos reais.
    4. Adiciona colunas de TotalOrçado, TotalAcumulado e Saldo por WBS.
    """
    try:
        # --- Leitura Robusta dos Arquivos ---
        # Leitura do arquivo de orçamento
        if orcamento_file.name.lower().endswith('.xls'):
            try:
                df_orcamento = pd.read_excel(orcamento_file, engine='xlrd')
            except Exception:
                orcamento_file.seek(0) # Retorna ao início do arquivo
                df_orcamento = pd.read_csv(orcamento_file, sep='\t', encoding='latin1', header=0)
        else:
            df_orcamento = pd.read_excel(orcamento_file)

        # Leitura do arquivo de gastos
        if gastos_file.name.lower().endswith('.xls'):
            try:
                df_gastos = pd.read_excel(gastos_file, engine='xlrd')
            except Exception:
                gastos_file.seek(0) # Retorna ao início do arquivo
                df_gastos = pd.read_csv(gastos_file, sep='\t', encoding='latin1', header=0)
        else:
            df_gastos = pd.read_excel(gastos_file)

        # --- Limpeza de Colunas "Unnamed" ---
        df_orcamento = df_orcamento.loc[:, ~df_orcamento.columns.str.startswith('Unnamed')]
        df_gastos = df_gastos.loc[:, ~df_gastos.columns.str.startswith('Unnamed')]
        
        # --- ETAPA 1: ATUALIZAÇÃO DOS GASTOS ---
        gastos_agrupados = df_gastos.groupby(['Fiscal Year', 'WBS Element', 'Period'])['Vbl. value/Obj. curr'].sum().reset_index()
        df_resultado = df_orcamento.copy()

        for periodo_num in range(periodo_inicial, periodo_final + 1):
            target_col_orcamento = f'Period {periodo_num}'
            if target_col_orcamento not in df_resultado.columns:
                continue
            gastos_do_periodo = gastos_agrupados[(gastos_agrupados['Period'] == periodo_num) & (gastos_agrupados['Fiscal Year'] == ano)]
            coluna_gasto_temp = 'gasto_temp_para_atualizar'
            if not gastos_do_periodo.empty:
                gastos_do_periodo = gastos_do_periodo.rename(columns={'Vbl. value/Obj. curr': coluna_gasto_temp})
                df_resultado = pd.merge(df_resultado, gastos_do_periodo[['WBS Element', 'Fiscal Year', coluna_gasto_temp]], on=['WBS Element', 'Fiscal Year'], how='left')
            else:
                df_resultado[coluna_gasto_temp] = pd.NA
            condicao_ano = df_resultado['Fiscal Year'] == ano
            valores_atualizados = df_resultado.loc[condicao_ano, coluna_gasto_temp].fillna(0).round(0)
            df_resultado.loc[condicao_ano, target_col_orcamento] = valores_atualizados
            df_resultado.drop(columns=[coluna_gasto_temp], inplace=True)

        # --- ETAPA 2: CÁLCULOS DE TOTAL E SALDO ---
        colunas_periodo = sorted([col for col in df_resultado.columns if re.match(r'^Period \d+$', col)])
        if not colunas_periodo:
            st.warning("Nenhuma coluna no formato 'Period X' foi encontrada para calcular o total.")
            return df_resultado
        df_resultado['OrcamentoRealPorLinha'] = df_resultado[colunas_periodo].sum(axis=1)
        df_resultado['Total Acumulado'] = df_resultado.groupby('WBS Element')['OrcamentoRealPorLinha'].transform('sum')
        if 'Total' not in df_resultado.columns:
            st.error("Erro: A coluna 'Total' é necessária para o cálculo do Saldo, mas não foi encontrada no arquivo de Orçamento.")
            return None
        df_resultado['Total Orcado'] = df_resultado.groupby('WBS Element')['Total'].transform('sum')
        df_resultado['Saldo'] = df_resultado['Total Orcado'] - df_resultado['Total Acumulado']
        max_ano_por_wbs = df_resultado.groupby('WBS Element')['Fiscal Year'].transform('max')
        mask_nao_e_max_ano = df_resultado['Fiscal Year'] < max_ano_por_wbs
        df_resultado.loc[mask_nao_e_max_ano, ['Total Orcado', 'Total Acumulado', 'Saldo']] = pd.NA
        df_resultado.drop(columns=['OrcamentoRealPorLinha'], inplace=True)
        return df_resultado

    except Exception as e:
        # Em vez de messagebox, mostramos o erro na própria página
        st.error(f"Ocorreu um erro inesperado durante o processamento: {str(e)}")
        return None

# =================================================================================
# INTERFACE DO SITE (Streamlit)
# =================================================================================

st.set_page_config(layout="wide") # Deixa a página mais larga
st.title("Analisador de Orçamento e Gastos 📊")
st.markdown("Faça o upload dos arquivos de Orçamento e Gastos, defina os parâmetros e execute a análise.")

# --- Área de Upload ---
col1, col2 = st.columns(2)
with col1:
    orcamento_file = st.file_uploader("1. Arquivo de Orçamento (.xls, .xlsx)", type=["xls", "xlsx"])
with col2:
    gastos_file = st.file_uploader("2. Arquivo de Gastos - Smart List (.xls, .xlsx)", type=["xls", "xlsx"])

# --- Parâmetros ---
st.subheader("3. Parâmetros de Análise")
st.markdown("Selecione os parâmetros que deseja equalizar, por exemplo, Ano 2026, do Período 1 ao Período 3.")
col_ano, col_p_ini, col_p_fim = st.columns(3)
with col_ano:
    ano_val = st.number_input("Ano para atualizar", step=1, format="%d", value=2026)
with col_p_ini:
    p_ini_val = st.number_input("Período Inicial", min_value=1, max_value=12, step=1, value=1)
with col_p_fim:
    p_fim_val = st.number_input("Período Final", min_value=1, max_value=12, step=1, value=12)

st.divider() # Uma linha divisória

# --- Botão de Execução ---
if st.button("Executar Análise", type="primary", use_container_width=True):
    if orcamento_file is not None and gastos_file is not None:
        if p_fim_val < p_ini_val:
            st.error("Erro: O Período Final deve ser maior ou igual ao Período Inicial.")
        else:
            with st.spinner("Processando... Isso pode levar alguns segundos."):
                # --- Chama a Lógica ---
                df_final = processar_dados_streamlit(
                    orcamento_file,
                    gastos_file,
                    ano_val,
                    p_ini_val,
                    p_fim_val
                )

            # --- Oferece o Download ---
            if df_final is not None:
                st.success("Processamento concluído com sucesso!")
                
                # Prepara o arquivo Excel para download em memória
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, sheet_name='Relatorio_Atualizado', index=False)
                    # Adiciona a formatação de cores (exatamente como antes)
                    workbook  = writer.book
                    worksheet = writer.sheets['Relatorio_Atualizado']
                    negative_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    (max_row, max_col) = df_final.shape
                    if max_row > 0:
                        worksheet.conditional_format(1, 0, max_row, max_col - 1, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': negative_format})
                
                # Cria o botão de download
                st.download_button(
                    label="Clique aqui para baixar o Relatório Atualizado (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"Relatorio_Atualizado_{ano_val}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    else:

        st.warning("Por favor, faça o upload dos dois arquivos (Orçamento e Gastos) para continuar.")


