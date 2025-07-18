import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import threading

# ---------------- FUNÇÕES DE CÁLCULO ------------------

def calcula_isv_por_ano_periodo(df, nome_df):
    df = df.copy()
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce').dt.normalize()
    df['Mes'] = df['Data'].dt.month
    df['Ano'] = df['Data'].dt.year
    df['Periodo'] = np.where(df['Mes'].isin([10,11,12,1,2,3]), 'Umido', 'Seco')
    df['AnoRef'] = np.where(df['Mes'].isin([1,2,3]), df['Ano'] -1, df['Ano'])

    profundidades = [20, 40, 60]
    resultados = []
    grouped = df.groupby(['AnoRef', 'Periodo'])

    for (ano_ref, periodo), grupo in grouped:
        if grupo.empty:
            continue
        grupo = grupo.copy()
        grupo = grupo.sort_values('Data')
        
        for prof in profundidades:
            u_col = f'U{prof}'
            if u_col not in grupo.columns:
                continue

            umidade = grupo[['Data', u_col]].dropna().rename(columns={u_col: 'Umid'})
            umidade = umidade.sort_values('Data').reset_index(drop=True)

            umidade['Baixa'] = umidade['Umid'] < 0.360
            umidade['Grupo'] = (umidade['Baixa'] != umidade['Baixa'].shift()).cumsum()

            eventos = umidade.groupby(['Grupo', 'Baixa']).agg(
                Duracao=('Data', 'count')
            ).reset_index()

            eventos_validos = eventos[(eventos['Baixa']) & (eventos['Duracao'] >= 4)]

            nver = len(eventos_validos)
            dmax = eventos_validos['Duracao'].max() if nver > 0 else 0
            dver = eventos_validos['Duracao'].sum() if nver > 0 else 0

            isv = nver + ((1 / (1 + (0.0163 * dmax**2)**2.26))**0.17) - 0.001 * dver

            resultados.append({
                'Ano': ano_ref,
                'Periodo': periodo,
                'Origem': nome_df,
                'Profundidade': prof,
                'nver': nver,
                'dmax': dmax,
                'dver': dver,
                'ISV': isv
            })

    if not resultados:
        return None
    return pd.DataFrame(resultados)

def calcula_isv_varias_planilhas(planilhas):
    resultados = []
    for nome_df, df in planilhas.items():
        res = calcula_isv_por_ano_periodo(df, nome_df)
        if res is not None:
            resultados.append(res)
    if not resultados:
        return None
    return pd.concat(resultados, ignore_index=True)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")
    return output.getvalue()

# ---------------- INTERFACE STREAMLIT ------------------

st.set_page_config(page_title="Índice de Suficiência de Umidade (ISV)", layout="wide")

# Exibe o logo
st.image("IEMS_LOGO.png", width=80)  # ajuste o width como quiser

st.title("Calculadora do Índice de Suficiência de Umidade (ISV)")

uploaded_file = st.file_uploader("Envie seu arquivo Excel com várias abas", type=["xlsx"])

st.sidebar.header("Parâmetros para cálculo do ISV")

# Já não usamos mais "Tradicional" ou "Amplitude real" para o ISV
# Removemos as opções que não são mais necessárias.
st.sidebar.info("Este índice calcula o ISV com base em eventos de umidade baixa consecutivos.")

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    abas = xls.sheet_names
    st.write(f"Abas encontradas: {abas}")
    planilhas = {aba: xls.parse(aba) for aba in abas}

    # Calculando o ISV
    resultados_isv = calcula_isv_varias_planilhas(planilhas)

    if resultados_isv is not None and not resultados_isv.empty:
        st.subheader("Resultados do ISV")
        st.dataframe(resultados_isv)

        excel_data = to_excel(resultados_isv)
        st.download_button(
            label="📄 Baixar resultados em Excel",
            data=excel_data,
            file_name="ISV_resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Não foi possível calcular o ISV com os dados enviados.")
else:
    st.info("Faça upload do arquivo Excel para iniciar o cálculo.")

# --- Botão para encerrar o aplicativo ---
def fechar_app():
    def delayed_shutdown():
        import time
        time.sleep(1)
        os._exit(0)
    threading.Thread(target=delayed_shutdown).start()

if st.button("🚪 Encerrar aplicativo"):
    st.warning("Encerrando o aplicativo...")
    fechar_app()
