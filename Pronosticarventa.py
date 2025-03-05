import streamlit as st

import pandas as pd
from io import BytesIO
from openpyxl import Workbook
import matplotlib.pyplot as plt
import numpy as np
import xgboost as xgb
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Border, Side  # 必要なモジュールをインポート
import bleach  # bleachをインポート

# Secretsからパスワードを取得
PASSWORD = st.secrets["PASSWORD"]

# パスワード認証の処理
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "login_attempts" not in st.session_state:
    st.session_state.login_attempts = 0

def verificar_contraseña():
    contraseña_ingresada = st.text_input("Introduce la contraseña:", type="password")

    if st.button("Iniciar sesión"):
        if st.session_state.login_attempts >= 3:
            st.error("Has superado el número máximo de intentos. Acceso bloqueado.")
        elif contraseña_ingresada == PASSWORD:  # Secretsから取得したパスワードで認証
            st.session_state.authenticated = True
            st.success("¡Autenticación exitosa! Marque otra vez el botón 'Iniciar sesión'.")
        else:
            st.session_state.login_attempts += 1
            intentos_restantes = 3 - st.session_state.login_attempts
            st.error(f"Contraseña incorrecta. Te quedan {intentos_restantes} intento(s).")
        
        if st.session_state.login_attempts >= 3:
            st.error("Acceso bloqueado. Intenta más tarde.")

if st.session_state.authenticated:
    # 認証成功後に表示されるメインコンテンツ

    # 過去12か月の売上データの初期値
    ventas_iniciales = [21000, 17500, 18000, 18500, 25000, 21000, 19000, 22000, 23500, 19500, 21000, 23000]
    # 過去12か月のその他の特徴量（Touristasは2022年の数値、千人単位、Cruceristas（通過者）は含まない。家族送金は2003～23年の月間平均。単位100万ドル）
    turistas = [44, 51, 71, 86, 69, 81, 85, 75, 54, 63, 71, 96]
    remesas = [477, 488, 581, 572, 618, 623, 606, 633, 599, 636, 573, 641]
    
    st.write("### :blue[Pronóstico (estimación) de ventas en próximos 12 meses]")
    st.write("###### (Herramienta de Inteligencia Artificial por Modelo XGBoost, ajustado del método de los mínimos cuadrados, para sectores de comercio y turísmo)")
    st.write("###### :red[Esta herramienta estima las ventas en futuro próximo, mediante la información sobre las ventas realizadas en estos 12 meses, los datos climáticos de la ciudad (a seleccionar) y el monto de remesas familiares por mes, el número de visitantes exteriores al país. Será probable que el resultado de estimación no sea precisa, debido a la limitación de los datos de variables explicativas.]")
    
    # 各都市のデータ
    ciudades = {
        "Tegucigalpa": {
            "lluvias": [0.4, 0.5, 0.5, 2.7, 9.8, 13.3, 10.6, 12.3, 15.0, 11.1, 3.7, 1.3],
            "temperaturas": [20, 21, 22, 24, 24, 23, 22, 23, 22, 22, 21, 20],
        },
        "San Marcos de Colón": {
            "lluvias": [0.2, 0.3, 0.8, 2.4, 8.9, 11.7, 8.5, 11.0, 13.5, 10.9, 3.5, 1.0],
            "temperaturas": [21, 22, 23, 24, 23, 22, 22, 22, 21, 21, 21, 21],
        },
        "Choluteca": {
            "lluvias": [0.1, 0.2, 0.7, 2.3, 8.9, 11.7, 8.6, 11.0, 13.6, 10.7, 3.2, 0.9],
            "temperaturas": [29, 29, 30, 31, 29, 28, 29, 29, 27, 27, 28, 29],
        },
        "Santa Rosa de Copán": {
            "lluvias": [1.9, 1.8, 1.6, 3.0, 8.2, 13.2, 12.3, 12.3, 13.4, 9.9, 4.8, 2.9],
            "temperaturas": [18, 19, 20, 22, 22, 22, 21, 22, 21, 20, 19, 18],
        },
        "San Pedro Sula": {
            "lluvias": [5.9, 4.7, 3.8, 3.4, 6.4, 11.1, 11.3, 11.2, 11.5, 10.3, 8.5, 7.2],
            "temperaturas": [23, 24, 25, 27, 28, 28, 27, 27, 27, 26, 24, 24],
        },
    }
    # 月の選択肢
    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    
    st.write("##### :blue[Seleccione el mes actual y la ciudad cuyo clima es semejante al mismo de su lugar]")
    
    col1, col2 = st.columns(2)
    with col1:
        # 選択された月の初期値
        mes_actual = st.selectbox("Selecciona el mes actual", meses, index=7)
    
    with col2:
        # Select the city
        ciudad = st.selectbox("Selecciona la ciudad", list(ciudades.keys()))
    
    # Get the city's data
    lluvias = ciudades[ciudad]["lluvias"]
    temperaturas = ciudades[ciudad]["temperaturas"]
    
    # 月のインデックスを取得
    mes_index = meses.index(mes_actual)
    
    # ユーザーが売上データを入力
    st.write("##### :blue[Ingrese los datos de ventas de los últimos 12 meses]")
    
    # 各列に4か月分の売上データ入力フィールドを配置するための列の作成
    cols = st.columns(4)
    
    # 12か月前からの順序を保持し、各列に4か月分を表示
    for i in range(12):
        col_index = i // 3  # 0, 1, 2, 3 (4列)
        month_label = f"Hace {12 - i} meses ({meses[(mes_index - 12 + i) % 12]})"
        with cols[col_index]:
            ventas_iniciales[i] = st.number_input(month_label, value=ventas_iniciales[i], key=i)
    
    # データフレームの作成
    data = pd.DataFrame({
        'Ventas': ventas_iniciales,
        "Días de lluvias": lluvias[mes_index:] + lluvias[:mes_index],
        "Temperatura mínima del día": temperaturas[mes_index:] + temperaturas[:mes_index],
        'Visitantes exteriores al país': turistas[mes_index:] + turistas[:mes_index],
        "Remesas familiares": remesas[mes_index:] + remesas[:mes_index],
    })
    
    # 特徴量とターゲットの準備
    X = data[['Días de lluvias', 'Temperatura mínima del día', 'Visitantes exteriores al país', 'Remesas familiares']]
    y = data['Ventas']
    
    # データを訓練セットとテストセットに分割
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, shuffle=False)
    
    # XGBoostモデルの訓練
    model = xgb.XGBRegressor(objective='reg:squarederror', n_estimators=13)
    model.fit(X_train, y_train)
    
    # 12カ月先まで予測
    forecast_input = X.iloc[-1].values.reshape(1, -1)
    forecast = []
    for i in range(12):
        next_pred = model.predict(forecast_input)[0]
        forecast.append(next_pred)
        # 新しい特徴量の生成
        new_row = np.array([lluvias[(mes_index + i + 1) % 12], temperaturas[(mes_index + i + 1) % 12], turistas[(mes_index + i + 1) % 12], remesas[(mes_index + i + 1) % 12]]).reshape(1, -1)
        forecast_input = new_row
    
    forecast_df = pd.DataFrame(forecast, index=[f"{meses[(mes_index+i)%12]}" for i in range(12, 24)], columns=['Ventas'])
    forecast_df['Ventas'] = forecast_df['Ventas'].round(0).astype(int)  # 売上高を整数に丸める

    # 最小二乗法で傾きを計算
    from scipy.stats import linregress
    x = np.arange(len(ventas_iniciales))
    slope, intercept, _, _, _ = linregress(x, ventas_iniciales)

    # 傾きを加算して予測を修正
    forecast_df['Ventas'] = forecast_df['Ventas'] + slope * np.arange(1, 13)
    
    # 実績データと予測データの結合
    full_data = pd.concat([data, forecast_df])
    full_data.index = [f"Hace {12-i} meses ({meses[(mes_index-12+i)%12]})" for i in range(12)] + [meses[(mes_index+i)%12] for i in range(12, 24)]
    
    if st.button("Estimar (pronosticar) ventas futuras por la inteligencia artificial"):
    
        # グラフの表示
        st.subheader("Ventas realizadas y estimadas en los 24 meses")
        plt.figure(figsize=(12, 4))
        plt.plot(full_data.index[:12], full_data['Ventas'][:12], label='Ventas realizadas', color='blue', marker='o')
        plt.plot(full_data.index[12:], full_data['Ventas'][12:], label='Ventas estimadas', color='orange', marker='o')
        plt.xticks(rotation=45, ha='right')
        plt.legend(loc='upper left')
        plt.grid(True)
        plt.tight_layout()
        st.pyplot(plt)
    
        # 小数点以下を表示しない設定
        pd.options.display.float_format = '{:.0f}'.format   
        
        # 表の表示
        st.subheader("Datos de ventas realizadas y estimadas")
        st.write("Los datos de días de lluvia y otros indicadores no son exactamente del año pasado sino de los otros años de muestra.")
        resultados = pd.concat([data, forecast_df.round(0)])
        resultados.index = [f"Hace {12-i} meses ({meses[(mes_index-12+i)%12]})" for i in range(12)] + [meses[(mes_index+i)%12] for i in range(12, 24)]
        st.dataframe(resultados)
    
        # エクセルファイルのダウンロード
        st.subheader("Descargar Datos en Excel")
        def convert_df(df):
            return df.to_csv().encode('utf-8')
        csv = convert_df(resultados)
        st.download_button(label="Descargar datos en Excel como CSV", data=csv, file_name='prediccion_ventas.csv', mime='text/csv')

else:
    verificar_contraseña()
