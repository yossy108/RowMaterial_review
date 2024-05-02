import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import statistics
from io import BytesIO

#初期値 (ファイル読み込み前のエラー防止)
item_list=[]
ins_list=[]
df_uploaded = pd.DataFrame()
df_review = pd.DataFrame()

# タイトル
st.title("ABF原料管理図トレンド")

# 解析対象ファイルの読み込みボタン
uploaded_file = st.file_uploader("読み込み用ファイルを選択してください")
if uploaded_file is not None:
    df_uploaded = pd.read_csv(uploaded_file, encoding='cp932')
    df_review = df_uploaded
    # st.dataframe(df_uploaded.head())    # 削除
    
    # 受入日をdatetime型に変換
    df_uploaded["受入日"] = pd.to_datetime(df_uploaded["受入日"])

    # 空白列の削除
    df_uploaded = df_uploaded.dropna(axis=1, how="all")
    # st.dataframe(df_uploaded.head())

    # 品目名称の指定文字列の削除（プルダウンを見やすくする、同品目のものを統一化するため）
    specified_strings = ["球状ｼﾘｶ", "ｱﾃﾞｶｽﾀﾌﾞ", "ｱﾃﾞｶﾚｼﾞﾝ", "ｱﾄﾞﾏﾌｧｲﾝ", "ｴｽﾚｯｸ", "ｴﾎﾟﾄｰﾄ", "ｶﾙﾎﾞｼﾞﾗｲﾄ", "ｽﾀﾌｨﾛｲﾄﾞ", "ﾃｲｻﾝﾚｼﾞﾝ", "ﾌｪﾉﾗｲﾄ", "ﾏﾘｱﾘﾑ", "ﾗﾋﾞﾄﾙ", "ﾙｸｼﾃﾞｨｱ"]
    for string in specified_strings:
        df_uploaded["品目名称"] = df_uploaded["品目名称"].str.replace(string, "").str.strip()

    # ()の削除（同品目のものを統一化するため）
    df_uploaded["品目名称"] = df_uploaded["品目名称"].replace("\(.*\)","",regex=True).replace("（.*）","",regex=True)

    # 規格・管理値設定をしていない場合、nanにreplace。
    df_uploaded["USL"] = df_uploaded["USL"].replace(999999, np.nan)
    df_uploaded["LSL"] = df_uploaded["LSL"].replace(-99999, np.nan)
    df_uploaded["UCL"] = df_uploaded["UCL"].replace(999999, np.nan)
    df_uploaded["LCL"] = df_uploaded["LCL"].replace(-99999, np.nan)
    # st.dataframe(df_uploaded.head())  # 削除
    
    # 品目名称をitem_listに格納
    item_list = sorted(df_uploaded["品目名称"].unique())
    st.write('ファイルは正常に読み込まれました')

# サイドバーに表示する項目
st.sidebar.write("## 条件指定")
item_selected = st.sidebar.selectbox("品目名称", item_list)
if item_selected is not None:
    # 品目名称で絞り込みした品目にデータフレームを絞り込み
    df_uploaded = df_uploaded[df_uploaded["品目名称"]==item_selected]
    # 検査項目をins_listに格納
    ins_list = sorted(df_uploaded["検査項目"].unique())
    ins_selected = st.sidebar.selectbox("検査項目", ins_list)

# UCLCRの算出式を定義
# 欠損値(nan)の場合、σ=0の場合の場合は"-"とする
def calc_UCLCR(data, UCL):
    if not np.isnan(UCL):
        mean_value = statistics.mean(data)
        std_value = statistics.stdev(data)
        if std_value != 0:
            UCLCR = ((mean_value+3*std_value)-UCL)/std_value
            UCLCR = round(UCLCR,2)
        else: UCLCR = "-"
    else: UCLCR = "-"
    return UCLCR

# LCLCRの算出式を定義
# 欠損値(nan)の場合、σ=0の場合の場合は"-"とする
def calc_LCLCR(data, LCL):
    if not np.isnan(LCL):
        mean_value = statistics.mean(data)
        std_value = statistics.stdev(data)
        if std_value != 0:
            LCLCR = (LCL-(mean_value-3*std_value))/std_value
            LCLCR = round(LCLCR,2)
        else: LCLCR = "-"
    else: LCLCR = "-"
    return LCLCR

# Cpkの算出式を定義
# USL/LSLともに欠損値(nan)の場合、ともに0の場合(外観など)は"-"とする
def calc_Cpk(data, USL, LSL):
    mean_value = statistics.mean(data)
    std_value = statistics.stdev(data)
    if np.isnan(USL) and np.isnan(LSL):
        Cpk = "-"
    elif [USL, LSL] == [0, 0]:
        Cpk = "-"
    elif not np.isnan(USL) and not np.isnan(LSL):
        Cpk = min((USL-mean_value)/(3*std_value), (mean_value-LSL)/(3*std_value))
    elif np.isnan(USL):
        Cpk = (mean_value-LSL)/(3*std_value)
    elif np.isnan(LSL):
        Cpk = (USL-mean_value)/(3*std_value)
    return Cpk

if uploaded_file is not None:

    if ins_selected is not None:
        # 検査項目の絞り込み後、受入日順に並べ替え
        df_uploaded = df_uploaded[df_uploaded["検査項目"]==ins_selected]
        df_uploaded = df_uploaded.sort_values(by="受入日", ascending=True)
        # st.dataframe(df_uploaded.head())    # 削除

        # ロット重複削除（最終受入ロットを残す）
        df_uploaded = df_uploaded.drop_duplicates(subset="ロット", keep="last", ignore_index=True)
        
        # 現行の規格値・管理値の抽出、CLCRの計算（規格値・管理値は最終受入ロットから抽出）
        cur_USL = df_uploaded["USL"].tail(1).iloc[0]
        cur_LSL = df_uploaded["LSL"].tail(1).iloc[0]
        cur_UCL = df_uploaded["UCL"].tail(1).iloc[0]
        cur_LCL = df_uploaded["LCL"].tail(1).iloc[0]
        cur_UCLCR = calc_UCLCR(df_uploaded["測定値"], cur_UCL)
        cur_LCLCR = calc_LCLCR(df_uploaded["測定値"], cur_LCL)

        # 現行実績の統計量を計算
        num = len(df_uploaded["測定値"])
        avg = statistics.mean(df_uploaded["測定値"])
        std = statistics.stdev(df_uploaded["測定値"])
        max_value = max(df_uploaded["測定値"])
        min_value = min(df_uploaded["測定値"])
        max_value = max(df_uploaded["測定値"])
        min_value = min(df_uploaded["測定値"])
        Cpk = calc_Cpk(df_uploaded["測定値"], cur_USL, cur_LSL)

        # 新しい管理値案、CLCRを計算
        # 初期値はAve±3σに設定のため、CLCRの初期値は0となる
        new_UCL = st.sidebar.number_input("UCL案（初期値 = Ave + 3σ）", value = avg+3*std)
        new_LCL = st.sidebar.number_input("LCL案（初期値 = Ave - 3σ）", value = avg-3*std)
        new_UCLCR = calc_UCLCR(df_uploaded["測定値"], new_UCL)
        new_LCLCR = calc_LCLCR(df_uploaded["測定値"], new_LCL)

        # # 新しい管理値の表示桁数を現行の管理値に合わせる処理（見た目の問題のみ）
        # cur_UCL_decimal_count = len(str(cur_UCL).split(".")[1]) if "." in str(cur_UCL) else 0
        # formatted_new_UCL = "{:.{}f}".format(new_UCL, cur_UCL_decimal_count)
        # cur_LCL_decimal_count = len(str(cur_LCL).split(".")[1]) if "." in str(cur_LCL) else 0
        # formatted_new_LCL =  "{:.{}f}".format(new_LCL, cur_LCL_decimal_count)

        # USL/LSLのmax/minを算出
        max_USL = max(df_uploaded["USL"])
        min_LSL = min(df_uploaded["LSL"])

        # グラフY軸の上下限を計算
        # 実績、新管理値、新規格値の中で最大値・最小値をY軸上下限とする
        y_axis_upper = max(max_value, new_UCL, max_USL)
        y_axis_lower = min(min_value, new_LCL, min_LSL)

        # 新しい管理線をグラフに入れるため、データフレームに追加
        df_uploaded["new_UCL"] = new_UCL
        df_uploaded["new_LCL"] = new_LCL

        # # グラフ上下限設定の計算が正しくできているかの確認（一応）
        # df_uploaded["y_axis_upper"] = y_axis_upper
        # df_uploaded["y_axis_lower"] = y_axis_lower
        
        # st.dataframe(df_uploaded.tail(5))    # 削除
        
        # Y軸の単位設定
        if not df_uploaded["単位"].tail(1).empty:
            unit = df_uploaded["単位"].tail(1).iloc[0]
        else:
            unit = "-"

        # ベース作成（毎回記載すると冗長化するため）
        base = alt.Chart(df_uploaded).encode(alt.X("ロット:N", title="Lot", sort=None))

        # グラフ作成
        chart = base.mark_line(point=True).encode(
            alt.Y("測定値:Q", title=item_selected + "_" + ins_selected + "   " + f"[{unit}]", scale=alt.Scale(domain=[y_axis_lower, y_axis_upper]))
            )

        # 新旧CL
        cur_UCL_line = base.mark_line(color="green").encode(
            alt.Y("UCL:Q", scale=alt.Scale(domain=[y_axis_lower, y_axis_upper])))
        cur_LCL_line = base.mark_line(color="green").encode(
            alt.Y("LCL:Q", scale=alt.Scale(domain=[y_axis_lower, y_axis_upper])))
        new_UCL_line = base.mark_line(color="red", strokeDash=[2,2]).encode(
            alt.Y("new_UCL:Q", scale=alt.Scale(domain=[y_axis_lower, y_axis_upper])))
        new_LCL_line = base.mark_line(color="red", strokeDash=[2,2]).encode(
            alt.Y("new_LCL:Q", scale=alt.Scale(domain=[y_axis_lower, y_axis_upper])))

        # # SL
        # cur_USL_line = base.mark_line(color="orange", strokeDash=[2,2]).encode(
        #     alt.Y("USL:Q", scale=alt.Scale(domain=[y_axis_lower, y_axis_upper])))
        # cur_LSL_line = base.mark_line(color="orange", strokeDash=[2,2]).encode(
        #     alt.Y("LSL:Q", scale=alt.Scale(domain=[y_axis_lower, y_axis_upper])))

        # データを重ねる
        # SLなし
        layer = alt.layer(chart, cur_UCL_line, cur_LCL_line, new_UCL_line, new_LCL_line)

        # # SLあり
        # layer = alt.layer(chart, cur_UCL_line, cur_LCL_line, new_UCL_line, new_LCL_line, cur_USL_line, cur_LSL_line)

        # タイトル、グラフの表示
        st.write('## トレンドチャート')
        st.altair_chart(layer, use_container_width=True)

        # メトリクスの表示、データ取り込み前の初期値は - としておく
        col1, col2, col3, col4 = st.columns(4)
        if item_selected is None:
            col1.metric('N', '-')
            col1.metric('Average', '-')
            col1.metric('Std.', '-')
            col1.metric("Cpk", "-")
            col2.metric("Currend USL", "-")
            col2.metric("Current LSL", "-")
            col3.metric('Current UCL', '-')
            col3.metric('Current LCL', '-')
            col3.metric('Current UCLCR', '-')
            col3.metric('Current LCLCR', '-')
            col4.metric('New UCL', '-')
            col4.metric('New LCL', '-')
            col4.metric('New UCLCR', '-')
            col4.metric('New LCLCR', '-')
        
        else:
            col1.metric('N', num)
            col1.metric('Average', round(avg,2))
            col1.metric('σ', round(std,2))
            
            # Cpk="-"の場合、"-"とする
            # 現行のUCL/LCLが設定されていない場合、Cpk="-"になるようにdefで定義されている
            if Cpk == "-":
                col1.metric("Cpk", "-")
            else:
                col1.metric("Cpk", round(Cpk, 2))
            
            # 現行のUSL/LSLが設定されていない（欠損値, nan）場合は"-"とする
            if np.isnan(cur_USL):
                col2.metric("Currend USL", "-")
            else:
                col2.metric("Currend USL", cur_USL)
            
            if np.isnan(cur_LSL):
                col2.metric("Currend LSL", "-")
            else:
                col2.metric("Currend LSL", cur_LSL)

            # 現行のUCL/LCLが設定されていない（欠損値, nan）場合は"-"とする
            if np.isnan(cur_UCL):
                col3.metric('Current UCL', "-")
            else:
                col3.metric('Current UCL', cur_UCL)
            
            if np.isnan(cur_LCL):
                col3.metric('Current LCL', "-")
            else:
                col3.metric('Current LCL', cur_LCL)
            
            if np.isnan(cur_UCL) or cur_UCLCR == "-":
                col3.metric('Current UCLCR', "-")
            else:
                col3.metric('Current UCLCR', round(cur_UCLCR,2))
            
            if np.isnan(cur_LCL) or cur_LCLCR == "-":
                col3.metric('Current LCLCR', "-")
            else:
                col3.metric('Current LCLCR', round(cur_LCLCR,2))
            
            # Average, σともに"0"の項目は"-"とする（例：外観等の数値項目）
            if avg == 0 and std == 0:
                col4.metric('New UCL', "-", help="初期値はAve + 3σ")
                col4.metric('New LCL', "-", help="初期値はAve - 3σ")
            else:
                col4.metric('New UCL', round(new_UCL,2), help="初期値はAve + 3σ")
                col4.metric('New LCL', round(new_LCL,2), help="初期値はAve - 3σ")
            
            # 新しいUCLCR/LCLCR="-"、つまりσ=0でCLCRが計算できない場合、"-"とする
            # σ=0の時は、分母が0となるためCLCR="-"となるようにdefで定義している
            if new_UCLCR == "-":
                col4.metric('New UCLCR', "-", help="初期値は0.0")
            else:
                col4.metric('New UCLCR', round(new_UCLCR,2), help="初期値は0.0")
            
            if new_LCLCR == "-":
                col4.metric('New LCLCR', "-", help="初期値は0.0")
            else:
                col4.metric('New LCLCR', round(new_LCLCR,2), help="初期値は0.0")
            
        st.write('##### 直近10ロットデータ')
        st.dataframe(df_uploaded.tail(10))

    st.sidebar.write("## 全原料集計データ")
    if st.sidebar.button("All Summary"):
        # df_review = pd.read_csv(uploaded_file, encoding='cp932')

        st.write('## 全原料集計データ')

        # 全て空白(how="all")になっている列の削除
        df_review = df_review.dropna(axis=1, how="all")
        df_review["受入日"] = pd.to_datetime(df_review["受入日"])

        # 指定文字列の削除
        specified_strings = ["球状ｼﾘｶ", "ｱﾃﾞｶｽﾀﾌﾞ", "ｱﾃﾞｶﾚｼﾞﾝ", "ｱﾄﾞﾏﾌｧｲﾝ", "ｴｽﾚｯｸ", "ｴﾎﾟﾄｰﾄ", "ｶﾙﾎﾞｼﾞﾗｲﾄ", "ｽﾀﾌｨﾛｲﾄﾞ", "ﾃｲｻﾝﾚｼﾞﾝ", "ﾌｪﾉﾗｲﾄ", "ﾏﾘｱﾘﾑ", "ﾗﾋﾞﾄﾙ", "ﾙｸｼﾃﾞｨｱ"]
        for string in specified_strings:
            df_review["品目名称"] = df_review["品目名称"].str.replace(string, "").str.strip()

        # ()の削除
        df_review["品目名称"] = df_review["品目名称"].replace("\(.*\)","",regex=True).replace("（.*）","",regex=True)

        # ["測定値"]の列に欠損値(nan)がある場合は行ごと削除
        # 数値計算において影響はないはず。ただし、N数に影響してくるか？（同品目でもnanのある項目のみN数が少なくなる）
        df_review = df_review.dropna(subset = ["測定値"])

        # 日付順に並べ替え→重複削除
        df_review = df_review.sort_values(by="受入日", ascending=True)
        df_review = df_review.drop_duplicates(subset=["ロット", "検査項目"], keep="last", ignore_index=True) 

        # LCLCRの算出式を定義
        def calc_LCLCR(data, LCL):
            if LCL != -99999:
                mean_value = statistics.mean(data)
                std_value = statistics.stdev(data)
                if std_value != 0:
                    LCLCR = (LCL-(mean_value-3*std_value))/std_value
                    LCLCR = round(LCLCR,2)
                else: LCLCR = "-"
            else: LCLCR = "-"
            return LCLCR

        # ULCRの算出式を定義
        def calc_UCLCR(data, UCL):
            if UCL != 999999:
                mean_value = statistics.mean(data)
                std_value = statistics.stdev(data)
                if std_value != 0:
                    UCLCR = ((mean_value+3*std_value)-UCL)/std_value
                    UCLCR = round(UCLCR,2)
                else: UCLCR = "-"
            else: UCLCR = "-"
            return UCLCR

        # Cpkの算出式を定義
        def calc_Cpk(data, LSL, USL):
            mean_value = statistics.mean(data)
            std_value = statistics.stdev(data)
            if [LSL, USL] == [-99999, 999999] or [LSL, USL] == [0, 0]:
                Cpk = "-"
            elif [LSL, USL] != [-99999, 999999]:
                Cpk = round(min((USL-mean_value)/(3*std_value), (mean_value-LSL)/(3*std_value)),2)
            elif USL == 999999:
                Cpk = round((mean_value-LSL)/(3*std_value),2)
            elif LSL == -99999:
                Cpk = round((USL-mean_value)/(3*std_value),2)
            return Cpk

        # 空のリストを作成
        dfs_summary = []

        for group_keys, group_df in df_review.groupby(["品目名称", "検査項目"]):
            n = len(group_df["測定値"])
            if n > 1:
                avg = statistics.mean(group_df["測定値"])
                std = statistics.stdev(group_df["測定値"])
                avg_m3s = avg - 3*std 
                avg_p3s = avg + 3*std
                avg_m4s = avg - 4*std
                avg_p4s = avg + 4*std
                min_value = min(group_df["測定値"])
                max_value = max(group_df["測定値"])
                LSL_value = group_df["LSL"].tail(1).iloc[0]
                USL_value = group_df["USL"].tail(1).iloc[0]
                LCL_value = group_df["LCL"].tail(1).iloc[0]
                UCL_value = group_df["UCL"].tail(1).iloc[0]
                LCLCR_value = calc_LCLCR(group_df["測定値"], LCL_value)
                UCLCR_value = calc_UCLCR(group_df["測定値"], UCL_value)
                Cpk_value = calc_Cpk(group_df["測定値"], LCL_value, UCL_value)
                
                dfs_summary.append(pd.DataFrame({
                    "品目名称": [group_keys[0]],
                    "検査項目": [group_keys[1]],
                    "N数": [n],
                    "Avg": [avg],
                    "σ": [std],
                    "Avg-3σ":[avg_m3s],
                    "Avg+3σ":[avg_p3s],
                    "Avg-4σ":[avg_m4s],
                    "Avg+4σ":[avg_p4s],
                    "min":[min_value],
                    "max":[max_value],
                    "LSL":[LSL_value],
                    "USL":[USL_value],
                    "LCL":[LCL_value],
                    "UCL":[UCL_value],
                    "LCLCR":[LCLCR_value],
                    "UCLCR":[UCLCR_value],
                    "Cpk":[Cpk_value],
                }))
                
            else:
                dfs_summary.append(pd.DataFrame({
                    "品目名称": [group_keys[0]],
                    "検査項目": [group_keys[1]],
                    "N数": [0],
                    "Avg": [0],
                    "σ": [0],
                    "Avg-3σ": [0],
                    "Avg+3σ": [0],
                    "Avg-4σ": [0],
                    "Avg+4σ": [0],
                    "min": [0],
                    "max": [0],
                    "LSL": [0],
                    "USL": [0],
                    "LCL": [0],
                    "UCL": [0],
                    "LCLCR": [0],
                    "UCLCR": [0],
                    "Cpk": [0],
                }))

        df_summary = pd.concat(dfs_summary, ignore_index=True)
        df_summary

        # 平均値、σが0の行は削除
        df_filtered = df_summary[(df_summary['Avg'] != 0) & (df_summary['σ'] != 0)]
        # N数30以上を抽出
        df_for_review = df_filtered[df_filtered["N数"] >= 30]

        # Excelファイルに書き込むためのExcelWriterを作成
        summary_data_xlsx = BytesIO()
        with pd.ExcelWriter(summary_data_xlsx, engine='xlsxwriter') as writer:
            # df_summaryを1つ目のシートに書き込む
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            # df_reviewを2つ目のシートに書き込む
            df_for_review.to_excel(writer, sheet_name='Review(Lot≧30)', index=False)
        # Excelファイルを保存
        writer.save()
        out_xlsx = summary_data_xlsx.getvalue()

        st.download_button(label="Download All Summary", data=out_xlsx, file_name="summary_data.xlsx")
