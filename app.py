import streamlit as st
import requests
import openpyxl
import webbrowser
from io import BytesIO
import base64
from streamlit.components.v1 import html

# --- コアロジック関数群（変更なし） ---
def open_in_Maps(origin, optimized_segments):
    """最適化されたルートをブラウザで開く"""
    if not optimized_segments:
        return

    start_point = optimized_segments[0]['from']
    waypoints = [segment['to'] for segment in optimized_segments[:-1]]
    final_destination = optimized_segments[-1]['to']

    origin_encoded = requests.utils.quote(start_point)
    waypoints_encoded = requests.utils.quote('|'.join(waypoints))
    final_destination_encoded = requests.utils.quote(final_destination)
    
    url = (
        "https://www.google.com/maps/embed/v1/directions?"
        f"key={st.secrets['Maps_API_KEY']}&"
        f"origin={origin_encoded}&"
        f"destination={final_destination_encoded}&"
        f"waypoints={waypoints_encoded}"
    )
    return url

def read_addresses_from_excel(file_content):
    """Excelファイルから住所を読み込む"""
    try:
        workbook = openpyxl.load_workbook(file_content)
        sheet = workbook.active
        addresses = []
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and len(cell.value) > 5:
                    if any(keyword in cell.value for keyword in ['都', '道', '府', '県', '市', '区', '町', '丁目', '番地']):
                        addresses.append(cell.value)
        return addresses
    except Exception as e:
        st.error(f"ファイルの読み込みに失敗しました: {e}")
        return None

def get_optimized_route_data(api_key, origin, destinations):
    """Google Maps Directions APIでルートを最適化して情報を取得する"""
    if not destinations:
        return None
        
    waypoints_str = '|'.join(destinations)
    url = (
        "https://maps.googleapis.com/maps/api/directions/json?"
        f"origin={origin}&"
        f"destination={origin}&"
        f"waypoints=optimize:true|{waypoints_str}&"
        f"key={api_key}"
    )

    try:
        response = requests.get(url)
        data = response.json()

        if data['status'] == 'OK':
            route = data['routes'][0]
            legs = route['legs']
            
            total_distance = sum(leg['distance']['value'] for leg in legs) / 1000
            total_duration = sum(leg['duration']['value'] for leg in legs) / 60
            
            segments = []
            for leg in legs:
                segments.append({
                    'from': leg['start_address'],
                    'to': leg['end_address'],
                    'distance': round(leg['distance']['value'] / 100) / 10,
                    'time': round(leg['duration']['value'] / 60)
                })
            
            return {
                'total_distance': round(total_distance * 10) / 10,
                'total_time': round(total_duration),
                'segments': segments,
                'Maps_result': route
            }
        else:
            st.error(f"ルートの計算に失敗しました: {data['status']}")
            return None
    except requests.exceptions.RequestException as e:
        st.error(f"APIリクエストエラー: {e}")
        return None

# --- Streamlit UI構築 ---
st.set_page_config(
    page_title="Google Maps ルート最適化",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("Google Maps ルート最適化")
st.markdown("Google Maps Directions APIを使って、複数の目的地を巡回する最適なルートを計算します。")

# --- セッションステートの初期化 ---
if 'destinations' not in st.session_state:
    st.session_state.destinations = []

if 'optimized_route_data' not in st.session_state:
    st.session_state.optimized_route_data = None

if 'map_url' not in st.session_state:
    st.session_state.map_url = None
    
if 'new_dest_input' not in st.session_state:
    st.session_state.new_dest_input = ""

# --- 新しい関数 ---
def add_destination():
    new_dest = st.session_state.new_dest_input
    if new_dest:
        st.session_state.destinations.append(new_dest)
        st.session_state.new_dest_input = "" # 入力欄をクリア
        st.success(f"'{new_dest}' をリストに追加しました。")

def clear_route_data():
    st.session_state.destinations = []
    st.session_state.optimized_route_data = None
    st.session_state.map_url = None

try:
    api_key = st.secrets["Maps_API_KEY"]
except KeyError:
    st.error("APIキーが設定されていません。")
    st.stop()

# --- UIコンポーネント ---
with st.sidebar:
    st.header("設定")
    st.write("APIキーは安全な方法で読み込まれています。")
    start_location = st.text_input(
        "出発地",
        "〒062-0912 北海道札幌市豊平区水車町６丁目３−１"
    )

    st.header("目的地リスト")
    
    # 目的地の手動追加
    st.text_input("新しい目的地を追加", key="new_dest_input")
    st.button("追加", on_click=add_destination)

    # Excelファイルから読み込み
    uploaded_file = st.file_uploader("Excelファイルから住所を読み込む", type=["xlsx", "xls"])
    if uploaded_file:
        file_content = BytesIO(uploaded_file.getvalue())
        addresses_from_file = read_addresses_from_excel(file_content)
        if addresses_from_file:
            if len(addresses_from_file) > 23:
                st.warning("Excelから読み込んだ住所が23件を超えています。")
                st.session_state.addresses_to_select = addresses_from_file
            else:
                # 修正点: リストを置き換えるのではなく、追加する
                st.session_state.destinations.extend(addresses_from_file)
                st.success(f"{len(addresses_from_file)}件の住所をリストに追加しました。")
                st.rerun()

    # 目的地リストの表示と削除
    if st.session_state.destinations:
        st.subheader("現在の目的地")
        for i, dest in enumerate(st.session_state.destinations):
            col1, col2 = st.columns([0.8, 0.2])
            with col1:
                st.write(f"{i+1}. {dest}")
            with col2:
                if st.button("削除", key=f"del_{i}"):
                    st.session_state.destinations.pop(i)
                    st.rerun()

    # Excelから23件選択するUI（条件付き表示）
    if 'addresses_to_select' in st.session_state and st.session_state.addresses_to_select:
        with st.expander("読み込んだ住所から選択 (最大23件)", expanded=True):
            selected_addresses = st.multiselect(
                "選択してください",
                st.session_state.addresses_to_select
            )
            if len(selected_addresses) > 23:
                st.warning("23件までしか選択できません。")
            
            if st.button("選択を確定"):
                if len(selected_addresses) <= 23:
                    # 修正点: リストを置き換えるのではなく、追加する
                    st.session_state.destinations.extend(selected_addresses)
                    st.session_state.addresses_to_select = None
                    st.success(f"{len(selected_addresses)}件の目的地を選択しました。")
                    st.rerun()
                else:
                    st.error("23件以内で選択してください。")

# --- ルートクリアボタン ---
st.button("ルートをクリア", on_click=clear_route_data)

# メインコンテンツ
st.header("ルート計算")

if st.button("🚗 ルート最適化"):
    if not start_location or not st.session_state.destinations:
        st.error("出発地、目的地をすべて入力してください。")
    else:
        with st.spinner("ルートを最適化中..."):
            route_data = get_optimized_route_data(api_key, start_location, st.session_state.destinations)
            st.session_state.optimized_route_data = route_data
        
        if st.session_state.optimized_route_data:
            st.success("✅ ルート最適化が完了しました。")
            st.session_state.map_url = open_in_Maps(start_location, st.session_state.optimized_route_data['segments'])
            st.rerun()

# 結果表示
if st.session_state.optimized_route_data:
    info = st.session_state.optimized_route_data
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("最適化されたルート概要")
        st.metric("総走行距離", f"{info['total_distance']} km")
        st.metric("総運転時間", f"{info['total_time']} 分")
        st.markdown("---")
        st.subheader("ルート詳細")
        for i, segment in enumerate(info['segments']):
            st.write(f"**{i+1}. {segment['from']}** → **{segment['to']}**")
            st.caption(f"距離: {segment['distance']} km, 時間: {segment['time']} 分")
        
        if st.button("🌍 新しいタブで開く"):
            webbrowser.open_new_tab(st.session_state.map_url)

    if st.session_state.map_url:
        with col2:
            st.subheader("埋め込み地図")
            st.warning("※一部のブラウザのセキュリティ設定により、地図が表示されない場合があります。")
            html_code = f"""
            <iframe
              width="100%"
              height="500"
              frameborder="0" style="border:0"
              src="{st.session_state.map_url}"
              allowfullscreen>
            </iframe>
            """
            html(html_code, height=500)
