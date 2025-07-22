import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import threading

# ---------------- FUNÇÃO ISV ------------------

def calcular_ISV_por_profundidade(dados, coluna_umidade, umid_limite=0.360):
    dados["Data"] = pd.to_datetime(dados["Data"], errors='coerce').dt.normalize()
    dados[coluna_umidade] = dados[coluna_umidade].replace(0, np.nan)

    dados_diarios = (
        dados.groupby("Data")[coluna_umidade]
        .mean()
        .reset_index(name="Umedia")
    )
    dados_diarios["mes"] = dados_diarios["Data"].dt.month
    dados_diarios["ano_ciclo"] = np.where(
        dados_diarios["mes"].isin([1, 2, 3]),
        dados_diarios["Data"].dt.year - 1,
        dados_diarios["Data"].dt.year
    )
    dados_diarios["periodo"] = np.where(
        dados_diarios["mes"].isin([10, 11, 12, 1, 2, 3]),
        "umido",
        "seco"
    )

    resultados = []

    for (ano_ciclo, periodo), grupo in dados_diarios.groupby(["ano_ciclo", "periodo"]):
        grupo = grupo.sort_values("Data").copy()
        grupo["abaixo_limite"] = (grupo["Umedia"] < umid_limite).astype(int)

        # Run-length encoding
        run_values = grupo["abaixo_limite"].values
        diffs = np.diff(np.concatenate(([0], run_values, [0])))
        run_starts = np.where(diffs == 1)[0]
        run_ends = np.where(diffs == -1)[0]
        comprimentos = run_ends - run_starts
        valores = run_values[run_starts]

        eventos = pd.DataFrame({"comprimento": comprimentos, "valor": valores})
        eventos_veranico = eventos[(eventos["valor"] == 1) & (eventos["comprimento"] >= 4)]

        nver = len(eventos_veranico)
        dver = eventos_veranico["comprimento"].sum()
        dmax = eventos_veranico["comprimento"].max() if nver > 0 else 0

        ISV = nver + ((1 / (1 + (0.0163 * dmax ** 2) ** 2.26)) ** 0.17) - 0.001 * dver

        resultados.append({
            "ano_ciclo": ano_ciclo,
            "periodo": periodo,
            "nver": nver,
            "dver": dver,
            "dmax": dmax,
            "ISV": ISV,
            "profundidade": coluna_umidade
        })

    return pd.DataFrame(resultados)

def calcula_ISV_em_planilha(df, nome_df, profundidades=["U20", "U40", "U60"]):
    resultados = []
    for coluna in profundidades:
        if coluna in df.columns:
            res = calcular_ISV_por_profundidade(df, coluna)
            res["Origem"] = nome_df
            resultados.append(res)
    return pd.concat(resultados, ignore_index=True) if resultados else None

def calcula_ISV_varias_planilhas(planilhas):
    resultados = []
    for nome, df in planilhas.items():
        res = calcula_ISV_em_planilha(df, nome)
        if res is not None:
            resultados.append(res)
    return pd.concat(resultados, ignore_index=True) if resultados else None

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="ISV_Resultados")
    return output.getvalue()

# ---------------- INTERFACE STREAMLIT ------------------

st.set_page_config(page_title="Cálculo do ISV", layout="wide")
st.image("IEMS_LOGO.png", width=80)
st.title("Calculadora de ISV (Índice de Seca ou Veranico)")

uploaded_file = st.file_uploader("Envie seu arquivo Excel com várias abas", type=["xlsx"])

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    abas = xls.sheet_names
    st.write(f"Abas encontradas: {abas}")
    planilhas = {aba: xls.parse(aba) for aba in abas}

    resultados_isv = calcula_ISV_varias_planilhas(planilhas)

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

# Botão para encerrar o app
def fechar_app():
    import time
    time.sleep(1)
    os._exit(0)

if st.button("🚪 Encerrar aplicativo"):
    st.warning("Encerrando o aplicativo...")
    threading.Thread(target=fechar_app).start()
