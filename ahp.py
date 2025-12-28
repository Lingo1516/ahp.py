import streamlit as st
import pandas as pd
import numpy as np
import io

# 設定網頁標題
st.set_page_config(page_title="AHP 研究數據分析平台", layout="wide")

st.title("🏆 AHP 論文數據分析系統")
st.markdown("### 支援 Excel 多專家整合 • 自動矩陣運算")

# --- 數學運算核心函式 ---

def calculate_ahp(matrix):
    """計算單一矩陣的 AHP 權重與 CR"""
    n = matrix.shape[0]
    # 行加總
    col_sums = matrix.sum(axis=0)
    # 正規化
    normalized_matrix = matrix / col_sums
    # 算權重 (列平均)
    weights = normalized_matrix.mean(axis=1)
    
    # 算 CR
    # Lambda Max = Sum(行總和 * 權重)
    lambda_max = np.dot(col_sums, weights)
    ci = (lambda_max - n) / (n - 1)
    
    # RI 表 (擴充到 n=15)
    ri_table = {1:0, 2:0, 3:0.58, 4:0.90, 5:1.12, 6:1.24, 7:1.32, 8:1.41, 9:1.45, 10:1.49, 11:1.51, 12:1.48, 13:1.56, 14:1.57, 15:1.59}
    ri = ri_table.get(n, 1.49)
    cr = ci / ri if n > 2 else 0
    
    return weights, cr, ci

def geometric_mean_matrix(matrices):
    """計算多個矩陣的幾何平均"""
    # matrices 是一個 list of numpy arrays
    stack = np.array(matrices)
    #沿著第一個軸 (專家數) 算乘積，再開 n 次方根
    prod = np.prod(stack, axis=0)
    geo_mean = np.power(prod, 1/len(matrices))
    return geo_mean

def generate_excel_template(n_criteria, n_experts):
    """產生範例 Excel"""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    for i in range(n_experts):
        sheet_name = f'專家{i+1}'
        # 建立一個空的 DataFrame，只有標題
        cols = [f'指標{j+1}' for j in range(n_criteria)]
        df = pd.DataFrame(index=cols, columns=cols)
        
        # 寫入 Excel
        df.to_excel(writer, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # 寫入提示
        worksheet.write('A1', '請填寫黃色區域 (左下角會自動倒數)')
        
        # 加上黃色背景格式
        yellow_fmt = workbook.add_format({'bg_color': '#FFFF00', 'border': 1})
        
        # 寫入對角線 1 和公式
        # 注意：xlsxwriter 寫入是 (row, col) 從 0 開始
        # header 佔據了 row 0, index 佔據了 col 0
        start_row = 1
        start_col = 1
        
        for r in range(n_criteria):
            for c in range(n_criteria):
                cell_row = start_row + r
                cell_col = start_col + c
                
                # Excel 座標字串 (例如 B2)
                cell_ref =  xlsxwriter_utility.xl_rowcol_to_cell(cell_row, cell_col)
                
                if r == c:
                    worksheet.write(cell_row, cell_col, 1)
                elif r < c:
                    # 右上角 (使用者填寫區) - 預設填空或 1
                    worksheet.write(cell_row, cell_col, 1, yellow_fmt)
                else:
                    # 左下角 (公式區) = 1 / 對稱格
                    # 對稱格座標
                    target_row = start_row + c
                    target_col = start_col + r
                    target_ref = xlsxwriter_utility.xl_rowcol_to_cell(target_row, target_col)
                    worksheet.write_formula(cell_row, cell_col, f'=1/{target_ref}')

    writer.close()
    processed_data = output.getvalue()
    return processed_data

import xlsxwriter.utility as xlsxwriter_utility # 輔助計算座標

# --- 介面佈局 ---

st.sidebar.header("📥 步驟 1：下載範例檔")
criteria_count = st.sidebar.number_input("指標數量 (N)", min_value=3, max_value=15, value=4)
expert_count = st.sidebar.number_input("專家數量", min_value=1, max_value=20, value=3)

if st.sidebar.button("產生並下載 Excel 範例"):
    excel_data = generate_excel_template(criteria_count, expert_count)
    st.sidebar.download_button(
        label="點此下載 .xlsx 範例檔",
        data=excel_data,
        file_name=f"AHP_範例_{criteria_count}x{criteria_count}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.write("---")
st.header("📂 步驟 2：上傳分析")
uploaded_file = st.file_uploader("請上傳填寫好的 Excel 檔案", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 讀取所有 Sheet
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        valid_matrices = []
        expert_results = []
        
        st.write(f"偵測到 {len(sheet_names)} 位專家資料...")
        
        for sheet in sheet_names:
            # 讀取數據，不讀標題 (header=None)，之後再清理
            df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
            
            # 清理數據：只保留純數字的部分
            # 轉換為 numeric，無法轉的變 NaN，然後丟掉含有 NaN 的行列
            df_numeric = df.apply(pd.to_numeric, errors='coerce')
            
            # 找到最密集的數字區塊 (簡單做法：移除全空的行列)
            df_clean = df_numeric.dropna(how='all').dropna(axis=1, how='all')
            
            # 轉為 numpy array
            matrix = df_clean.values
            
            # 檢查是否為正方形
            rows, cols = matrix.shape
            if rows > 0 and rows == cols:
                weights, cr, ci = calculate_ahp(matrix)
                is_pass = cr < 0.1
                
                expert_results.append({
                    "專家": sheet,
                    "CR值": round(cr, 4),
                    "結果": "✅ 有效" if is_pass else "❌ 剔除",
                    "矩陣": matrix
                })
                
                if is_pass:
                    valid_matrices.append(matrix)
            else:
                st.warning(f"工作表 '{sheet}' 格式錯誤，無法讀取為正方形矩陣。")

        # 顯示個別專家結果
        if expert_results:
            st.subheader("1. 個別專家一致性檢定")
            results_df = pd.DataFrame(expert_results)
            st.dataframe(results_df[["專家", "CR值", "結果"]])
            
            # 顯示最終整合
            if valid_matrices:
                st.subheader("2. 群體決策整合結果 (幾何平均法)")
                
                final_matrix = geometric_mean_matrix(valid_matrices)
                final_weights, final_cr, final_ci = calculate_ahp(final_matrix)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("有效問卷數", f"{len(valid_matrices)} / {len(sheet_names)}")
                with col2:
                    st.metric("整合後 CR 值", f"{final_cr:.4f}", delta="合格" if final_cr < 0.1 else "不合格")
                
                # 權重排名表
                st.write("### 最終權重排名")
                rank_data = {
                    "指標": [f"指標 {i+1}" for i in range(len(final_weights))],
                    "權重": final_weights,
                    "百分比": [f"{w:.2%}" for w in final_weights]
                }
                rank_df = pd.DataFrame(rank_data).sort_values(by="權重", ascending=False).reset_index(drop=True)
                rank_df.index += 1 # 排名從 1 開始
                st.dataframe(rank_df)
                
                # 畫長條圖
                st.bar_chart(pd.Series(final_weights, index=rank_data["指標"]))
                
            else:
                st.error("沒有任何專家的 CR 值小於 0.1，無法進行整合。")
                
    except Exception as e:
        st.error(f"檔案讀取失敗：{e}")
