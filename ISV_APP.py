import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import threading

# ---------------- FUNÇÃO PRINCIPAL DE CÁLCULO DO ISV ----------------

def calcular_ISV_por_profundidade(dados, coluna_umidade, umid_limite=0.360, dias_evento=4):
    dados = dados.copy()
    dados["Data"] = pd.to_datetime(dados["Data"], errors='coerce').dt.normalize()
    dados[coluna_umidade] = dados[coluna_umidade].replace(0, np.nan)

    dados_diarios = dados.groupby("Data")[coluna_umidade].mean().reset_index(name="Umedia")
    dados_diarios["mes"] = dados_diarios["Data"].dt.month
    dados_diarios["ano_ciclo"] = np.where(dados_diarios["mes"].isin([1, 2, 3]),
                                          dados_diarios["Data"].dt.year - 1,
                                          dados_diarios["Data"].dt.year)
    dados_diarios["periodo"] = np.where(dados_diarios["mes"].isin([10, 11, 12, 1, 2, 3]),
                                        "umido", "seco")

    def calcular_ISV_grupo(df):
        df = df.sort_values("Data").copy()
        df["abaixo_limite"] = (df["Umedia"] < umid_limite).astype(int)

        rle = (df["abaixo_limite"] != df["abaixo_limite"].shift()).cumsum()
        eventos = df.groupby(rle).agg({
            "abaixo_limite": ["first", "count"]
        })
        eventos.columns = ["valor", "comprimento"]
        eventos_veranico = eventos[(eventos["valor"] == 1) & (eventos["comprimento"] >= dias_evento)]

        nver = len(eventos_veranico)
        dver = eventos_veranico["comprimento"].sum()
        dmax = eventos_veranico["comprimento"].max() if nver > 0 else 0

        ISV = nver + ((1 / (1 + (0.0163 * dmax**2)**2.26))**0.17) - 0.001 * dver

        return pd.DataFrame([{
            "nver": nver,
            "dmax": dmax,
            "dver": dver,
            "ISV": ISV
        }])

    resultado = dados_diarios.groupby(["ano_ciclo", "periodo"]).apply(calcular_ISV_grupo).reset_index(drop=True)
    resultado["profundidade"] = coluna_umidade
    return resultado

# ---------------- FUNÇÃO PARA VÁRIAS PLANILHAS ----------------

def calcula_isv_varias_planilhas(planilhas, umid_limite=0.360, dias_evento=4):
    resultados = []
    for nome_df, df in planilhas.items():
        for col in ["U20", "U40", "U60"]:
            if col in df.columns:
                res = calcular_ISV_por_profundidade(df, col, umid_limite=umid_limite, dias_evento=dias_evento)
                res["Origem"] = nome_df
                resultados.append(res)
    return pd.concat(resultados, ignore_index=True) if resultados else None

# ---------------- GERADOR DE ARQUIVO EXCEL ----------------

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados ISV")
    return output.getvalue()

# ---------------- INTERFACE STREAMLIT ----------------

st.set_page_config(page_title="ISV - Veranico", layout="wide")
st.title("Calculadora de ISV (Índice de Severidade do Veranico)")

st.markdown("Envie um arquivo `.xlsx` com **várias abas**, contendo as colunas de umidade (ex: `U20`, `U40`, `U60`) e datas.")

uploaded_file = st.file_uploader("📤 Enviar arquivo Excel com várias abas", type=["xlsx"])

# --------- Parâmetros ajustáveis ---------
st.sidebar.header("Parâmetros de Cálculo")

umid_limite = st.sidebar.slider("Umidade limite para veranico (m³/m³)", 0.100, 0.500, 0.360, step=0.005)
dias_evento = st.sidebar.slider("Mínimo de dias consecutivos para caracterizar evento", 4, 10, 4)

# --------- Processamento ---------
if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    abas = xls.sheet_names
    st.write(f"📄 Abas encontradas no arquivo: {abas}")
    planilhas = {aba: xls.parse(aba) for aba in abas}

    resultados_isv = calcula_isv_varias_planilhas(
        planilhas,
        umid_limite=umid_limite,
        dias_evento=dias_evento
    )

    if resultados_isv is not None and not resultados_isv.empty:
        st.subheader("📊 Resultados do ISV")
        st.dataframe(resultados_isv)

        excel_data = to_excel(resultados_isv)
        st.download_button(
            label="⬇️ Baixar resultados em Excel",
            data=excel_data,
            file_name="ISV_resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ Nenhum resultado foi gerado. Verifique as colunas de umidade e datas.")
else:
    st.info("📁 Faça upload de um arquivo Excel com os dados.")

# --------- Encerramento opcional ---------
def fechar_app():
    def delayed_shutdown():
        import time
        time.sleep(1)
        os._exit(0)
    threading.Thread(target=delayed_shutdown).start()

if st.button("🚪 Encerrar aplicativo"):
    st.warning("Encerrando o aplicativo...")
    fechar_app()
