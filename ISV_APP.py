import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import threading

# ---------------- FUNÇÕES DE CÁLCULO ------------------

# Função para agrupar por data e criar as colunas de mês, ano e período
def agrupar_por_data(df):
    df = df.copy()
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce').dt.normalize()  # Normaliza a data
    df['Mes'] = df['Data'].dt.month
    df['Ano'] = df['Data'].dt.year
    df['Periodo'] = np.where(df['Mes'].isin([10, 11, 12, 1, 2, 3]), 'Umido', 'Seco')
    df['AnoRef'] = np.where(df['Mes'].isin([1, 2, 3]), df['Ano'] - 1, df['Ano'])
    return df

# Função que detecta eventos de baixa umidade
def detectar_eventos_baixa_umidade(df, umidade_limite=0.360, dias_consecutivos=4):
    eventos = []
    
    profundidades = [20, 40, 60]
    periodos = ['Seco', 'Umido']
    
    for prof in profundidades:
        u_col = f'U{prof}'
        
        # Verificar se a coluna existe
        if u_col not in df.columns:
            print(f"Coluna {u_col} não encontrada no DataFrame. Pulando essa profundidade.")
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

# Função que calcula o ISV com base nos eventos detectados
def calcular_isv(eventos):
    resultados = []
    
    for prof, periodo, df_eventos in eventos:
        if df_eventos.empty:
            continue
        
        nver = len(df_eventos['Consecutivo'].unique())  # Número de eventos
        dmax = df_eventos['Data'].max() - df_eventos['Data'].min()  # Duração do maior evento
        dver = len(df_eventos)  # Soma das durações dos eventos (número de dias com eventos)
        ano = df_eventos['AnoRef'].iloc[0]  # Ano de referência do evento (pode variar entre os eventos)

        # Calcular o ISV com a fórmula fornecida
        isv = nver + ((1 / (1 + (0.0163 * dmax.days**2)**2.26))**0.17) - 0.001 * dver
        resultados.append({
            'Ano': ano,  # Adicionando o ano
            'Profundidade': prof,
            'Periodo': periodo,
            'nver': nver,
            'dmax': dmax.days,  # Convertendo de timedelta para dias
            'dver': dver,
            'ISV': isv
        })
    
    return pd.DataFrame(resultados)

# Função que calcula o ISV por ano e período
def calcula_isv_por_ano_periodo(df, nome_df):
    df = agrupar_por_data(df)  # Aplica o agrupamento por data
    # Detecta eventos de baixa umidade
    eventos = detectar_eventos_baixa_umidade(df, umidade_limite=0.360, dias_consecutivos=4)
    
    # Calcula o ISV com base nos eventos
    resultado_isv = calcular_isv(eventos)
    
    if resultado_isv is not None and not resultado_isv.empty:
        resultado_isv['Origem'] = nome_df
        return resultado_isv
    else:
        return None

# Função para calcular o ISV por várias planilhas
def calcula_isv_varias_planilhas(planilhas):
    resultados = []
    for nome_df, df in planilhas.items():
        resultado_isv = calcula_isv_por_ano_periodo(df, nome_df)
        
        if resultado_isv is not None:
            resultados.append(resultado_isv)

    if resultados:
        return pd.concat(resultados, ignore_index=True)
    else:
        return None

# Função para exportar resultados para Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")
    return output.getvalue()

# ---------------- INTERFACE STREAMLIT ------------------

st.set_page_config(page_title="Índice Microclimático", layout="wide")

# Exibe o logo
st.image("IEMS_LOGO.png", width=80)  # ajuste o width como quiser

st.title("Calculadora de Índices Microclimáticos do Solo")

uploaded_file = st.file_uploader("Envie seu arquivo Excel com várias abas", type=["xlsx"])

st.sidebar.header("Parâmetros para cálculo do ISV")

# Parâmetros para o cálculo do ISV (se necessário)
umidade_limite = st.sidebar.slider("Limite de umidade para eventos", 0.0, 1.0, 0.360)
dias_consecutivos = st.sidebar.slider("Número de dias consecutivos", 1, 10, 4)

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    abas = xls.sheet_names
    st.write(f"Abas encontradas: {abas}")
    planilhas = {aba: xls.parse(aba) for aba in abas}

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
