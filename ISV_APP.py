import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import threading

# ---------------- FUNÇÕES DE CÁLCULO ------------------

def agrupar_por_data(df):
    """
    Agrupa os dados por data e remove a hora, garantindo que seja um formato de data sem hora.
    """
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce').dt.normalize()  # Normaliza a data
    return df

def detectar_eventos_baixa_umidade(df, umidade_limite=0.360, dias_consecutivos=4):
    """
    Detecta eventos de baixa umidade, considerando pelo menos 'dias_consecutivos' dias consecutivos.
    """
    eventos = []
    
    profundidades = [20, 40, 60]
    periodos = ['Seco', 'Umido']
    
    for prof in profundidades:
        u_col = f'U{prof}'
        if u_col not in df.columns:
            continue
        
        for periodo in periodos:
            df_periodo = df[df['Periodo'] == periodo]
            
            # Identificar os dias em que a umidade é abaixo do limite
            df_periodo['Evento'] = df_periodo[u_col] < umidade_limite
            
            # Marcar sequências de dias consecutivos
            df_periodo['Consecutivo'] = (df_periodo['Evento'] != df_periodo['Evento'].shift()).cumsum()

            # Filtrar eventos com pelo menos 4 dias consecutivos
            eventos_ano_periodo = df_periodo.groupby('Consecutivo').filter(
                lambda x: len(x) >= dias_consecutivos and x['Evento'].all()
            )
            
            # Armazenar eventos por profundidade e período
            eventos.append((prof, periodo, eventos_ano_periodo))

    return eventos

def calcular_isv(eventos):
    """
    Calcula o ISV com base nos eventos detectados.
    """
    resultados = []
    
    for prof, periodo, df_eventos in eventos:
        nver = len(df_eventos['Consecutivo'].unique())  # Número de eventos
        dmax = df_eventos['Data'].max() - df_eventos['Data'].min()  # Duração do maior evento
        dver = len(df_eventos)  # Soma das durações dos eventos (número de dias com eventos)
        
        # Calcular o ISV com a fórmula fornecida
        isv = nver + ((1 / (1 + (0.0163 * dmax.days**2)**2.26))**0.17) - 0.001 * dver
        resultados.append({
            'Profundidade': prof,
            'Periodo': periodo,
            'nver': nver,
            'dmax': dmax.days,  # Convertendo de timedelta para dias
            'dver': dver,
            'ISV': isv
        })
    
    return pd.DataFrame(resultados)

def calcular_isv_completo(df):
    """
    Função que integra todos os passos: agrupar por data, detecção de eventos e cálculo do ISV.
    """
    # Passo 1: Agrupar por data
    df = agrupar_por_data(df)
    
    # Passo 2: Determinar os eventos de baixa umidade (sequências de pelo menos 4 dias consecutivos)
    eventos = detectar_eventos_baixa_umidade(df)
    
    # Passo 3: Calcular o ISV
    resultado_isv = calcular_isv(eventos)
    
    return resultado_isv

def to_excel(df):
    """
    Converte o DataFrame para um arquivo Excel.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")
    return output.getvalue()

# ---------------- INTERFACE STREAMLIT ------------------

st.set_page_config(page_title="Cálculo do ISV", layout="wide")

# Exibe o logo
st.image("IEMS_LOGO.png", width=80)  # ajuste o width como quiser

st.title("Calculadora de ISV - Índice de Sequência de Baixa Umidade")

uploaded_file = st.file_uploader("Envie seu arquivo Excel com várias abas", type=["xlsx"])

st.sidebar.header("Parâmetros para cálculo")

umidade_limite = st.sidebar.slider("Limite da umidade (abaixo de qual valor?), padrão 0.360", 0.0, 1.0, 0.360)
dias_consecutivos = st.sidebar.slider("Número de dias consecutivos com baixa umidade", 4, 10, 4)

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    abas = xls.sheet_names
    st.write(f"Abas encontradas: {abas}")
    planilhas = {aba: xls.parse(aba) for aba in abas}

    # Processar os dados e calcular o ISV
    resultados_isv = []
    for nome_df, df in planilhas.items():
        res = calcular_isv_completo(df)
        if res is not None and not res.empty:
            res['Origem'] = nome_df
            resultados_isv.append(res)
    
    if resultados_isv:
        # Concatenar os resultados de todas as planilhas
        df_final_isv = pd.concat(resultados_isv, ignore_index=True)
        st.subheader("Resultados do ISV")
        st.dataframe(df_final_isv)

        # Botão para download
        excel_data = to_excel(df_final_isv)
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
