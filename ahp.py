import streamlit as st
import pandas as pd
import numpy as np

# --- 頁面設定 ---
st.set_page_config(page_title="AHP 層級分析系統 V6.0", layout="wide")

# --- 核心數學函式 ---
def repair_matrix(matrix):
    """
    修復矩陣 (單一專家)：
    1. 確保對角線為 1
    2. 確保右上角有值 (若無則補1)
    3. 自動計算左下角倒數 (這是關鍵！必須在幾何平均前做)
    """
    matrix = np.array(matrix, dtype=float)
    rows, cols = matrix.shape
    
    for i in range(rows):
        for j in range(cols):
            if i == j: 
                matrix[i, j] = 1.0
            elif i < j:
                # 右上角：如果讀到 0 或 NaN，預設補 1
                if matrix[i, j] == 0 or np.isnan(matrix[i, j]): 
                    matrix[i, j] = 1.0
                # 左下角：強制倒數
                if matrix[i, j] != 0:
                    matrix[j, i] = 1.0 / matrix[i, j]
                else:
                    matrix[j, i] = 1.0 
    return matrix

def calculate_ahp_weights(matrix):
    """只計算權重與 CR (不需再修復，因為進來前已經修復過了)"""
    n = matrix.shape[0]
    col_sums = matrix.sum(axis=0)
    with np.errstate(divide='ignore', invalid='ignore'):
        normalized_matrix = matrix / col_sums
    weights = normalized_matrix.mean(axis=1)
    
    lambda_max = np.dot(col_sums, weights)
    ci = (lambda_max - n) / (n - 1) if n > 1 else 0
    ri_table = {1:0, 2:0, 3:0.58, 4:0.90, 5:1.12, 6:1.24, 7:1.32, 8:1.41, 9:1.45, 10:1.49}
    ri = ri_table.get(n, 1.49)
    cr = ci / ri if n > 2 else 0
    return weights, cr

def geometric_mean_matrix(matrices):
    """多專家幾何平均"""
    if not matrices: return None
    stack = np.array(matrices)
    # 這裡因為傳進來的 matrices 都已經被 repair 過了，所以不會有 0
    prod = np.prod(stack, axis=0)
    geo_mean = np.power(prod, 1/len(matrices))
    return geo_mean

# --- 主程式介面 ---

st.title("⚖️ AHP 層級分析系統 (V6.0 修正版)")
st.markdown("已修正：幾何平均運算邏輯、矩陣補 0 問題、Matplotlib 錯誤。")

tab1, tab2 = st.tabs(["Step 1: 計算局部權重", "Step 2: 整合全球權重"])

# === Tab 1: 權重計算器 ===
with tab1:
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.info("💡 操作提示：上傳 Excel 後，系統會自動修補矩陣並進行幾何平均整合。")
        uploaded_file = st.file_uploader("上傳 Excel 檔", type=['xlsx', 'xls'])
        
        st.write("---")
        st.markdown("**✂️ 矩陣裁切設定**")
        manual_n = st.number_input("強制設定指標數量 (N)", min_value=0, max_value=15, value=0, help="若出現 8 個指標但您只有 3 個，請輸入 3。")

    with col2:
        if uploaded_file is not None:
            try:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                valid_matrices = []
                
                st.write(f"📄 偵測到 {len(sheet_names)} 位專家資料")

                for sheet in sheet_names:
                    # 1. 讀取
                    df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                    df = df.apply(pd.to_numeric, errors='coerce')
                    df_clean = df.dropna(how='all').dropna(axis=1, how='all')
                    raw_matrix = df_clean.values
                    
                    # 2. 裁切
                    if manual_n > 0:
                        if raw_matrix.shape[0] >= manual_n and raw_matrix.shape[1] >= manual_n:
                            raw_matrix = raw_matrix[:manual_n, :manual_n]
                    
                    rows, cols = raw_matrix.shape
                    
                    if rows == cols and rows > 1:
                        # 3. 【關鍵修正】先修復矩陣 (填補 0)，才加入列表
                        repaired_matrix = repair_matrix(raw_matrix)
                        valid_matrices.append(repaired_matrix)
                    else:
                        st.warning(f"⚠️ 工作表 {sheet} 格式異常，已略過。")

                if valid_matrices:
                    # 4. 幾何平均整合
                    final_matrix = geometric_mean_matrix(valid_matrices)
                    
                    # 5. 計算最終權重
                    weights, cr = calculate_ahp_weights(final_matrix)
                    
                    st.success("✅ 計算完成！")
                    
                    # 顯示整合後的矩陣 (確認用)
                    with st.expander("👀 查看整合後的矩陣 (幾何平均)", expanded=False):
                        st.dataframe(pd.DataFrame(final_matrix))

                    # 結果顯示
                    res_col1, res_col2 = st.columns(2)
                    with res_col1:
                        st.metric("整合後 CR 值", f"{cr:.4f}", delta="合格" if cr < 0.1 else "不一致", delta_color="inverse")
                    
                    # 表格
                    df_res = pd.DataFrame({
                        "指標": [f"指標 {i+1}" for i in range(len(weights))],
                        "權重": weights
                    })
                    
                    # 這裡使用安全的顯示方式，避免 Matplotlib 錯誤
                    try:
                        st.dataframe(df_res.style.format({"權重": "{:.2%}"}).background_gradient(cmap="Blues"))
                    except:
                        # 萬一還是缺套件，就顯示純文字表格
                        st.dataframe(df_res.style.format({"權重": "{:.2%}"}))
                    
                    st.caption("請複製此處權重，填入 Step 2 進行整合。")

                else:
                    st.error("無法讀取有效矩陣。")

            except Exception as e:
                st.error(f"發生錯誤：{e}")

# === Tab 2: 全球權重整合 ===
with tab2:
    st.markdown("### 🌍 全球權重計算表")
    if "grid_data" not in st.session_state:
        st.session_state.grid_data = pd.DataFrame(
            [{"構面": "構面A", "構面權重": 0.5, "準則": "準則A1", "準則局部權重": 0.6}]
        )

    edited_df = st.data_editor(st.session_state.grid_data, num_rows="dynamic", use_container_width=True)

    if st.button("計算最終排名"):
        res = edited_df.copy()
        res["構面權重"] = pd.to_numeric(res["構面權重"], errors='coerce').fillna(0)
        res["準則局部權重"] = pd.to_numeric(res["準則局部權重"], errors='coerce').fillna(0)
        res["全球權重"] = res["構面權重"] * res["準則局部權重"]
        res = res.sort_values("全球權重", ascending=False).reset_index(drop=True)
        st.dataframe(res.style.format({
            "構面權重": "{:.2%}", "準則局部權重": "{:.2%}", "全球權重": "{:.2%}"
        }))
