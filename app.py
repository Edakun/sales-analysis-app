import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from io import BytesIO

# アプリケーションのタイトルとスタイル設定
st.set_page_config(
    page_title="企業売上分析アプリ",
    page_icon="📊",
    layout="wide"
)

st.title("📊 企業売上分析アプリ")
st.markdown("### 昨年と今年の売上データを比較分析するツール")

# セッション状態の初期化
if 'last_year_data' not in st.session_state:
    st.session_state.last_year_data = None
if 'this_year_data' not in st.session_state:
    st.session_state.this_year_data = None
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'comparison_df' not in st.session_state:
    st.session_state.comparison_df = None

# サイドバーの設定
st.sidebar.title("データアップロード")
st.sidebar.markdown("### 手順")
st.sidebar.markdown("1. 昨年のデータをアップロード")
st.sidebar.markdown("2. 今年のデータをアップロード")
st.sidebar.markdown("3. 分析結果を確認")

# ファイルのアップロード機能
last_year_file = st.sidebar.file_uploader("昨年の売上データ（Excel形式）", type=['xlsx'])
this_year_file = st.sidebar.file_uploader("今年の売上データ（Excel形式）", type=['xlsx'])

# データ処理関数
def process_excel_data(file):
    """
    Excelファイルを読み込み、得意先ごとの受注額を集計します。
    """
    try:
        df = pd.read_excel(file)
        
        # 必要なカラムの確認
        required_columns = ['得意先', '受注額']
        optional_columns = ['担当者', '品名', '粗利益(B-L)']
        
        # 必須カラムの存在確認
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"必須カラム {', '.join(missing_columns)} がデータに見つかりません。")
            st.info("Excelファイルには少なくとも「得意先」と「受注額」のカラムが必要です。")
            return None
        
        # 任意カラムの存在確認と警告
        missing_optional = [col for col in optional_columns if col not in df.columns]
        if missing_optional:
            st.warning(f"一部のカラム {', '.join(missing_optional)} が見つかりません。基本的な分析のみ実行します。")
        
        # 数値型の確認と変換
        if df['受注額'].dtype == 'object':
            try:
                df['受注額'] = pd.to_numeric(df['受注額'], errors='coerce')
                df = df.dropna(subset=['受注額'])
                st.info("数値以外の受注額データは除外されました。")
            except:
                st.error("「受注額」カラムを数値に変換できません。データを確認してください。")
                return None
        
        # 得意先ごとの受注額を集計
        result_df = df.groupby('得意先')['受注額'].sum().reset_index()
        
        # 粗利益がある場合は追加
        if '粗利益(B-L)' in df.columns:
            if df['粗利益(B-L)'].dtype == 'object':
                try:
                    df['粗利益(B-L)'] = pd.to_numeric(df['粗利益(B-L)'], errors='coerce')
                except:
                    st.warning("「粗利益(B-L)」カラムを数値に変換できません。粗利益の分析は除外されます。")
            
            if df['粗利益(B-L)'].dtype != 'object':  # 変換に成功した場合
                profit_by_customer = df.groupby('得意先')['粗利益(B-L)'].sum().reset_index()
                result_df = pd.merge(result_df, profit_by_customer, on='得意先', how='left')
        
        # カラム名を標準化（売上分析関数との互換性のため）
        result_df = result_df.rename(columns={'得意先': '顧客名', '受注額': '売上金額'})
        
        return result_df
        
    except Exception as e:
        st.error(f"データの処理中にエラーが発生しました: {e}")
        return None

# 売上分析関数
def analyze_sales_data(last_year, this_year):
    """
    昨年と今年の売上データを比較し、増減を分析します。
    """
    # 顧客名で両方のデータをマージ
    merged_df = pd.merge(last_year, this_year, on='顧客名', how='outer', suffixes=('_前年', '_今年'))
    
    # NaN値を0に置き換え
    merged_df = merged_df.fillna(0)
    
    # 増減額と増減率を計算
    merged_df['増減額'] = merged_df['売上金額_今年'] - merged_df['売上金額_前年']
    
    # 0除算を避ける
    merged_df['増減率'] = np.where(
        merged_df['売上金額_前年'] != 0,
        (merged_df['増減額'] / merged_df['売上金額_前年']) * 100,
        np.inf  # 前年が0の場合は無限大とする
    )
    
    # 無限大の値を適切に表示するための処理
    merged_df['増減率'] = merged_df['増減率'].replace([np.inf, -np.inf], '新規/完全減少')
    
    # 数値型の列のみを整形
    for col in merged_df.select_dtypes(include=['float64', 'int64']).columns:
        if col != '増減率' or not isinstance(merged_df[col].iloc[0], str):
            merged_df[col] = merged_df[col].round(0).astype('int64')
    
    # 売上金額_今年の降順でソート
    merged_df = merged_df.sort_values(by='売上金額_今年', ascending=False)
    
    return merged_df

# ファイルが両方アップロードされた場合の処理
if last_year_file and this_year_file:
    # データ処理
    last_year_data = process_excel_data(last_year_file)
    this_year_data = process_excel_data(this_year_file)
    
    if last_year_data is not None and this_year_data is not None:
        st.session_state.last_year_data = last_year_data
        st.session_state.this_year_data = this_year_data
        
        # 売上分析の実行
        comparison_df = analyze_sales_data(last_year_data, this_year_data)
        st.session_state.comparison_df = comparison_df
        st.session_state.analysis_done = True
        
        st.success("データの分析が完了しました！")

# 分析結果の表示
if st.session_state.analysis_done and st.session_state.comparison_df is not None:
    st.markdown("## 売上分析結果")
    
    # フィルタリングオプション
    col1, col2 = st.columns(2)
    
    with col1:
        # 金額範囲フィルター
        min_amount = int(st.session_state.comparison_df['売上金額_今年'].min())
        max_amount = int(st.session_state.comparison_df['売上金額_今年'].max())
        selected_range = st.slider(
            "売上金額範囲（今年）",
            min_amount, max_amount, (min_amount, max_amount)
        )
    
    with col2:
        # 増減フィルター
        change_options = ['すべて表示', '増加のみ', '減少のみ']
        selected_change = st.selectbox("増減フィルター", change_options)
    
    # 検索フィルター
    search_term = st.text_input("顧客名で検索")
    
    # フィルタリング適用
    filtered_df = st.session_state.comparison_df.copy()
    
    # 金額範囲フィルター適用
    filtered_df = filtered_df[
        (filtered_df['売上金額_今年'] >= selected_range[0]) & 
        (filtered_df['売上金額_今年'] <= selected_range[1])
    ]
    
    # 増減フィルター適用
    if selected_change == '増加のみ':
        filtered_df = filtered_df[filtered_df['増減額'] > 0]
    elif selected_change == '減少のみ':
        filtered_df = filtered_df[filtered_df['増減額'] < 0]
    
    # 検索フィルター適用
    if search_term:
        filtered_df = filtered_df[filtered_df['顧客名'].str.contains(search_term, case=False, na=False)]
    
    # データテーブルの表示
    st.dataframe(filtered_df, use_container_width=True)
    
    # 集計情報の表示
    st.markdown("### 集計情報")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("前年総売上", f"{st.session_state.comparison_df['売上金額_前年'].sum():,.0f}円")
    
    with col2:
        st.metric("今年総売上", f"{st.session_state.comparison_df['売上金額_今年'].sum():,.0f}円")
    
    with col3:
        total_change = st.session_state.comparison_df['増減額'].sum()
        st.metric("総増減額", f"{total_change:,.0f}円")
    
    with col4:
        if st.session_state.comparison_df['売上金額_前年'].sum() > 0:
            change_rate = (total_change / st.session_state.comparison_df['売上金額_前年'].sum()) * 100
            st.metric("総増減率", f"{change_rate:.1f}%")
        else:
            st.metric("総増減率", "計算不能")
    
    # グラフ表示
    st.markdown("### 売上グラフ")
    
    # 上位N社を表示するオプション
    top_n = st.slider("表示する顧客数を選択", 5, 20, 10)
    
    # トップN社のデータを取得
    top_customers = filtered_df.head(top_n)
    
    # 棒グラフ
    fig = px.bar(
        top_customers,
        x='顧客名',
        y=['売上金額_前年', '売上金額_今年'],
        barmode='group',
        title=f'上位{top_n}社の売上比較',
        labels={'value': '売上金額', 'variable': '年度'},
        color_discrete_map={'売上金額_前年': '#1f77b4', '売上金額_今年': '#ff7f0e'}
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # 増減額グラフ
    fig2 = px.bar(
        top_customers,
        x='顧客名',
        y='増減額',
        title=f'上位{top_n}社の売上増減額',
        color='増減額',
        color_continuous_scale=['red', 'green'],
        labels={'増減額': '売上増減額'}
    )
    st.plotly_chart(fig2, use_container_width=True)
    
    # CSVダウンロード機能
    st.markdown("### データダウンロード")
    
    # CSVデータの準備
    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    
    # Excelデータの準備
    excel_buffer = BytesIO()
    filtered_df.to_excel(excel_buffer, index=False, engine='openpyxl')
    excel_data = excel_buffer.getvalue()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="CSV形式でダウンロード",
            data=csv,
            file_name="売上分析結果.csv",
            mime="text/csv",
        )
    
    with col2:
        st.download_button(
            label="Excel形式でダウンロード",
            data=excel_data,
            file_name="売上分析結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# アップロードファイルがない場合のガイダンス
if not last_year_file or not this_year_file:
    st.info("サイドバーから昨年と今年の売上データ（Excel形式）をアップロードしてください。")
    
    with st.expander("必要なExcelファイル形式について"):
        st.markdown("""
        ### 必要なExcelファイル形式
        
        アップロードするExcelファイルには、最低限以下のカラムが必要です：
        
        1. **得意先** - 顧客企業の名称
        2. **受注額** - 各顧客の受注額
        
        他の列が含まれていても問題ありませんが、上記の列名が必要です。
        
        #### サンプルデータ形式：
        
        | 得意先 | 受注額 | その他情報 |
        |--------|----------|------------|
        | A社    | 1000000  | ...        |
        | B社    | 2500000  | ...        |
        | C社    | 750000   | ...        |
        
        アップロードされたデータは、このアプリケーション内でのみ処理され、外部に送信されません。
        セッション終了後、データは自動的に削除されます。
        """)

# フッター
st.markdown("---")
st.markdown("© 2025 企業売上分析アプリ - データプライバシーを尊重しています。アップロードされたデータは外部に送信されません。")
