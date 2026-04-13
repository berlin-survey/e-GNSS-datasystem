import streamlit as st
import pandas as pd
import numpy as np
import io
import math
import itertools
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ================= 0. 系統初始化 =================
if 'temp_database' not in st.session_state:
    st.session_state.temp_database = []
if 'current_stage_data' not in st.session_state:
    st.session_state.current_stage_data = None
if 'current_stage_name' not in st.session_state:
    st.session_state.current_stage_name = ""
# 轉換相關
if 'trans_params' not in st.session_state:
    st.session_state.trans_params = None 
if 'trans_residuals' not in st.session_state:
    st.session_state.trans_residuals = None
if 'trans_rmse' not in st.session_state:
    st.session_state.trans_rmse = None
if 'trans_vv' not in st.session_state:
    st.session_state.trans_vv = None
if 'final_twd97_data' not in st.session_state:
    st.session_state.final_twd97_data = None
# 基線與實地檢測相關
if 'baseline_check_data' not in st.session_state:
    st.session_state.baseline_check_data = None

st.set_page_config(page_title="乾坤測繪 e-GNSS 雙測回檢核系統", layout="wide")
st.title("乾坤測繪 e-GNSS 雙測回檢核系統 🚀")

# ================= 1. 左側：參數設定與檔案 =================
col_left, col_right = st.columns([0.8, 1.2])

with col_left:
    st.header("⚙️ 參數與上傳")
    with st.expander("規範參數設定", expanded=True):
        st.write("--- 單測回品質 ---")
        err_plane = st.number_input("固定解平面誤差容許值 (m)", 0.02, format="%.3f")
        err_elev = st.number_input("固定解高程誤差容許值 (m)", 0.05, format="%.3f")
        min_pts = st.number_input("單測回有效筆數門檻", 180, help="測試時可設小一點")
        st.write("--- 雙測回比對 ---")
        diff_inst_limit = st.number_input("儀器高變換門檻 (m)", 0.10, format="%.3f")
        diff_plane_limit = st.number_input("平面較差容許值 (m)", 0.03, format="%.3f")
        diff_elev_limit = st.number_input("高程較差容許值 (m)", 0.05, format="%.3f")
        time_gap_limit = st.number_input("兩測回間隔時間 (分鐘)", 60)
        st.write("--- 轉換與檢核 ---")
        rmse_threshold = st.number_input("轉換中誤差(RMSE) 合格門檻 (m)", 0.03, format="%.3f")
        short_dist_limit = st.number_input("實地檢測短邊門檻 (m)", 100.0, help="小於此距離建議實地檢測")

    st.markdown("---")
    st.info("💡 支援批次上傳：您可以一次選取多個測點的 CSV/Excel 檔案。")
    files_round1 = st.file_uploader("📂 上傳 第 1 測回 觀測檔 (可多選)", type=['csv', 'xlsx'], accept_multiple_files=True)
    files_round2 = st.file_uploader("📂 上傳 第 2 測回 觀測檔 (可多選)", type=['csv', 'xlsx'], accept_multiple_files=True)
    
    if st.button("🎁 產生含時間欄位的標準測試檔"):
        data1 = "測點名稱,觀測時間,解算狀態,PDOP值,固定解平面誤差(m),固定解高程誤差(m),縱坐標_N(m),橫坐標_E(m),高程坐標_H(m),儀器高(m)\nA1-1,2023-10-25 09:00:00,Fixed,0.9,0.005,0.010,2540769.783,182723.175,62.032,1.50\nA1-2,2023-10-25 09:00:01,Fixed,0.9,0.005,0.010,2540769.785,182723.175,62.032,1.50\nA2-1,2023-10-25 09:10:00,Fixed,0.9,0.005,0.010,2540869.783,182823.175,63.032,1.50"
        data2 = "測點名稱,觀測時間,解算狀態,PDOP值,固定解平面誤差(m),固定解高程誤差(m),縱坐標_N(m),橫坐標_E(m),高程坐標_H(m),儀器高(m)\nA1-201,2023-10-25 10:30:00,Fixed,0.9,0.005,0.010,2540769.790,182723.180,62.040,1.65\nA2-201,2023-10-25 10:40:00,Fixed,0.9,0.005,0.010,2540869.790,182823.180,63.040,1.65"
        with open("第1測回_測試.csv", "w", encoding="utf-8-sig") as f: f.write(data1)
        with open("第2測回_測試.csv", "w", encoding="utf-8-sig") as f: f.write(data2)
        st.success("已產生測試檔。")

# ================= 2. 核心運算函數群 =================

def load_and_merge_files(file_list):
    df_list = []
    for file in file_list:
        try:
            if file.name.endswith('.csv'): df = pd.read_csv(file, encoding='utf-8-sig')
            else: df = pd.read_excel(file)
            df.columns = df.columns.str.replace('\n', '').str.replace('\r', '').str.replace('"', '').str.strip()
            df_list.append(df)
        except Exception as e: return None, f"檔案 {file.name} 讀取失敗: {str(e)}"
    if not df_list: return pd.DataFrame(), ""
    return pd.concat(df_list, ignore_index=True), ""

def process_single_round(file_list, round_name):
    log_text = []
    df, err_msg = load_and_merge_files(file_list)
    if err_msg: return None, [err_msg]
    if df.empty: return None, ["❌ 未讀取到任何資料"]

    try:
        required_cols = ['測點名稱', '解算狀態', '固定解平面誤差(m)', '固定解高程誤差(m)', '觀測時間']
        if not all(col in df.columns for col in required_cols): return None, [f"❌ 欄位缺失！需包含：{required_cols}"]

        mask = ((df['解算狀態'].astype(str).str.strip().str.upper() == 'FIXED') & (df['固定解平面誤差(m)'] <= err_plane) & (df['固定解高程誤差(m)'] <= err_elev))
        df_passed = df[mask]
        log_text.append(f"📊 **{round_name}**: 共 {len(df)} 筆 -> **過濾後**: {len(df_passed)} 筆")

        df_clean = df_passed.copy()
        df_clean['主測站'] = df_clean['測點名稱'].astype(str).apply(lambda x: x.split('-')[0] if '-' in x else x)
        
        # 修正時間解析：確保清除前後空白，並交由 Pandas 聰明解析各種 AM/PM 格式
        df_clean['觀測時間'] = df_clean['觀測時間'].astype(str).str.strip()
        df_clean['觀測時間_dt'] = pd.to_datetime(df_clean['觀測時間'], errors='coerce')

        final_results = []
        for station, group in df_clean.groupby('主測站'):
            mN, sN = group['縱坐標_N(m)'].mean(), group['縱坐標_N(m)'].std(ddof=1)
            mE, sE = group['橫坐標_E(m)'].mean(), group['橫坐標_E(m)'].std(ddof=1)
            mH, sH = group['高程坐標_H(m)'].mean(), group['高程坐標_H(m)'].std(ddof=1)
            sN, sE, sH = (0 if pd.isna(x) else x for x in [sN, sE, sH])
            
            mask_3s = ((abs(group['縱坐標_N(m)'] - mN) <= 3 * sN) & (abs(group['橫坐標_E(m)'] - mE) <= 3 * sE) & (abs(group['高程坐標_H(m)'] - mH) <= 3 * sH))
            valid_group = group[mask_3s]
            
            valid_pts = len(valid_group)
            total_pts = len(group)
            
            if valid_pts >= min_pts:
                # 🛠️ 修正時間平均的計算方式：直接使用 Pandas 原生 mean() 函數，完全避開整數與浮點數造成的時間軸錯亂
                valid_times = valid_group['觀測時間_dt'].dropna()
                mean_time = valid_times.mean() if not valid_times.empty else None
                if pd.isna(mean_time): 
                    mean_time = None
                
                # 計算該測回有效筆數之平均與中誤差 (為報表6-2準備)
                val_mN = valid_group['縱坐標_N(m)'].mean()
                val_mE = valid_group['橫坐標_E(m)'].mean()
                val_mH = valid_group['高程坐標_H(m)'].mean()
                
                val_sN = valid_group['縱坐標_N(m)'].std(ddof=1) if valid_pts > 1 else 0
                val_sE = valid_group['橫坐標_E(m)'].std(ddof=1) if valid_pts > 1 else 0
                val_sH = valid_group['高程坐標_H(m)'].std(ddof=1) if valid_pts > 1 else 0
                
                ratio = valid_pts / total_pts if total_pts > 0 else 0
                
                final_results.append({
                    '測點名稱': station, 
                    '測回別': round_name, 
                    '有效筆數': valid_pts, 
                    '總計點數': total_pts,
                    '使用比率': ratio,
                    '平均時間': mean_time, 
                    'N': val_mN, 
                    'E': val_mE, 
                    'H': val_mH, 
                    'sN': 0 if pd.isna(val_sN) else val_sN,
                    'sE': 0 if pd.isna(val_sE) else val_sE,
                    'sH': 0 if pd.isna(val_sH) else val_sH,
                    '儀器高': valid_group['儀器高(m)'].mean()
                })
                log_text.append(f"  ✅ {station}: 合格")
            else: log_text.append(f"  ❌ {station}: 剔除 (有效筆數 {valid_pts})")
        return final_results, log_text
    except Exception as e: return None, [f"處理錯誤: {str(e)}"]

def calc_dist_azimuth(n1, e1, n2, e2):
    dn = n2 - n1; de = e2 - e1
    dist = math.sqrt(dn**2 + de**2)
    az_rad = math.atan2(de, dn)
    az_deg = (math.degrees(az_rad) + 360) % 360
    return dist, az_deg

def deg_to_dmmss(deg):
    d = int(deg); m_full = (deg - d) * 60; m = int(m_full); s = (m_full - m) * 60
    return d + m/100 + s/10000

# --- Excel 輔助函數 ---
def setup_excel_style(ws, headers, row_idx=1):
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ws.append(headers) if row_idx == 1 else None 
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.value = header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.font = Font(bold=True)
        cell.border = border

def adjust_col_width(ws):
    for i, col in enumerate(ws.columns, start=1):
        col_letter = get_column_letter(i)
        max_len = 0
        for cell in col:
            if cell.value:
                try:
                    l = len(str(cell.value))
                    if any(ord(c) > 127 for c in str(cell.value)): l *= 1.5
                    if l > max_len: max_len = l
                except: pass
        ws.column_dimensions[col_letter].width = max(10, max_len + 2)

# --- 報表產出 ---
def generate_report_6_final(data, limits):
    wb = Workbook()
    ws = wb.active
    ws.title = "精度檢核總表"
    ws.merge_cells('A1:O1'); ws['A1'] = "e-GNSS 各階段精度檢核報表"; ws['A1'].font = Font(size=16, bold=True); ws['A1'].alignment = Alignment(horizontal='center')
    headers = ["測點名稱", "測回別", "觀測時間", "有效筆數", "單測回檢核", "縱坐標_N(m)", "橫坐標_E(m)", "高程坐標_H(m)", "儀器高(m)", "儀器高變換(m)", "平面較差(m)", "高程較差(m)", "時間間隔(分)", "雙測回比對", "備註"]
    setup_excel_style(ws, headers, row_idx=2)
    
    df = pd.DataFrame(data)
    stations = sorted(df['測點名稱'].unique())
    pass_font = Font(color="008000", bold=True); fail_font = Font(color="FF0000", bold=True)
    curr = 3
    for stn in stations:
        d = df[df['測點名稱'] == stn]
        r1 = d[d['測回別'] == '第 1 測回'].iloc[0].to_dict() if not d[d['測回別'] == '第 1 測回'].empty else None
        r2 = d[d['測回別'] == '第 2 測回'].iloc[0].to_dict() if not d[d['測回別'] == '第 2 測回'].empty else None
        
        if r1: 
            t1 = r1.get('平均時間').strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(r1.get('平均時間')) else ""
            ws.append([stn, "第 1 測回", t1, r1['有效筆數'], "✅合格", r1['N'], r1['E'], r1['H'], r1['儀器高'], "", "", "", "", "", ""])
            ws.cell(row=curr, column=5).font = pass_font; curr += 1
        if r2:
            comp = ""; vals = ["", "", "", ""]; notes = ""
            if r1:
                di, dp, de = abs(r1['儀器高']-r2['儀器高']), ((r1['N']-r2['N'])**2 + (r1['E']-r2['E'])**2)**0.5, abs(r1['H']-r2['H'])
                dt = abs((pd.to_datetime(r2['平均時間'])-pd.to_datetime(r1['平均時間'])).total_seconds())/60.0 if r1.get('平均時間') and r2.get('平均時間') else 0
                ok = (di>limits['diff_inst']) and (dp<=limits['diff_plane']) and (de<=limits['diff_elev']) and (dt>=limits['time_gap'])
                comp = "✅合格" if ok else "❌失敗"
                vals = [di, dp, de, dt]
                if not ok: notes = "檢核未過"
            t2 = r2.get('平均時間').strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(r2.get('平均時間')) else ""
            ws.append([stn, "第 2 測回", t2, r2['有效筆數'], "✅合格", r2['N'], r2['E'], r2['H'], r2['儀器高'], vals[0], vals[1], vals[2], vals[3], comp, notes])
            ws.cell(row=curr, column=5).font = pass_font
            ws.cell(row=curr, column=14).font = pass_font if "合格" in comp else fail_font
            curr += 1
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

# 🔥 報表 6-2 成果檢核表(測繪中心版)
def generate_report_6_2_center(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "成果檢核表(測繪中心版)"
    ws.merge_cells('A1:O1')
    ws['A1'] = "e-GNSS 即時動態定位坐標成果檢核表"
    ws['A1'].font = Font(size=16, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center')
    
    headers = ["點號", "N", "E", "h", "N_中誤差", "E_中誤差", "h_中誤差", "計算筆數", "總計點數", "觀測量使用比率", "平面較差", "高程較差", "平均N", "平均E", "平均h"]
    setup_excel_style(ws, headers, row_idx=2)
    
    df = pd.DataFrame(data)
    if df.empty: return io.BytesIO()
    stations = sorted(df['測點名稱'].unique())
    
    curr = 3
    for stn in stations:
        d = df[df['測點名稱'] == stn]
        r1 = d[d['測回別'] == '第 1 測回'].iloc[0].to_dict() if not d[d['測回別'] == '第 1 測回'].empty else None
        r2 = d[d['測回別'] == '第 2 測回'].iloc[0].to_dict() if not d[d['測回別'] == '第 2 測回'].empty else None
        
        avg_N = d['N'].mean()
        avg_E = d['E'].mean()
        avg_H = d['H'].mean()
        
        if r1:
            ws.append([
                f"{stn}-1", 
                round(r1['N'], 3), round(r1['E'], 3), round(r1['H'], 3),
                round(r1.get('sN', 0), 3), round(r1.get('sE', 0), 3), round(r1.get('sH', 0), 3),
                r1.get('有效筆數', ''), r1.get('總計點數', r1.get('有效筆數', '')), f"{r1.get('使用比率', 1)*100:.2f}%",
                "", "", "", "", ""
            ])
            curr += 1
            
        if r2:
            dp = ""
            dh = ""
            mN, mE, mH = "", "", ""
            if r1:
                dp = round(math.sqrt((r1['N'] - r2['N'])**2 + (r1['E'] - r2['E'])**2), 3)
                dh = round(abs(r1['H'] - r2['H']), 3)
                mN, mE, mH = round(avg_N, 3), round(avg_E, 3), round(avg_H, 3)
            else:
                mN, mE, mH = round(avg_N, 3), round(avg_E, 3), round(avg_H, 3)
                
            ws.append([
                f"{stn}-2", 
                round(r2['N'], 3), round(r2['E'], 3), round(r2['H'], 3),
                round(r2.get('sN', 0), 3), round(r2.get('sE', 0), 3), round(r2.get('sH', 0), 3),
                r2.get('有效筆數', ''), r2.get('總計點數', r2.get('有效筆數', '')), f"{r2.get('使用比率', 1)*100:.2f}%",
                dp, dh, mN, mE, mH
            ])
            curr += 1
            
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

def generate_report_2_coord(data):
    wb = Workbook(); ws = wb.active; ws.title = "坐標成果表"
    ws.merge_cells('A1:E1'); ws['A1'] = "e-GNSS 點位坐標成果表"; ws['A1'].font = Font(size=16, bold=True); ws['A1'].alignment = Alignment(horizontal='center')
    setup_excel_style(ws, ["測點名稱", "縱坐標_N(m)", "橫坐標_E(m)", "高程坐標_H(m)", "備註"], row_idx=2)
    df = pd.DataFrame(data)
    for stn in sorted(df['測點名稱'].unique()):
        g = df[df['測點名稱'] == stn]
        ws.append([stn, round(g['N'].mean(),3), round(g['E'].mean(),3), round(g['H'].mean(),3), "平均值"])
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

def generate_report_3_twd97(data):
    wb = Workbook(); ws = wb.active; ws.title = "TWD97成果表"
    ws.merge_cells('A1:E1'); ws['A1'] = "TWD97 點位坐標成果表 (強制附合)"; ws['A1'].font = Font(size=16, bold=True); ws['A1'].alignment = Alignment(horizontal='center')
    setup_excel_style(ws, ["測點名稱", "縱坐標_N(TWD97)", "橫坐標_E(TWD97)", "高程坐標_H", "備註"], row_idx=2)
    for row in data:
        ws.append([row['測點名稱'], round(row['N_TWD97'], 3), round(row['E_TWD97'], 3), round(row['H'], 3), "參數轉換+殘差分配"])
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

def generate_report_4_baseline(baseline_data):
    wb = Workbook(); ws = wb.active; ws.title = "基線比較表"
    ws.merge_cells('A1:K1'); ws['A1'] = "全組合基線比較表 (e-GNSS vs TWD97成果)"; ws['A1'].font = Font(size=16, bold=True); ws['A1'].alignment = Alignment(horizontal='center')
    headers = ["起點", "終點", "距離(e-GNSS)", "距離(TWD97)", "距離較差(m)", "相對精度(1/P)", "距離判定", "方位角(e-GNSS)", "方位角(TWD97)", "方位角差(秒)", "方位判定"]
    setup_excel_style(ws, headers, row_idx=2)
    for row in baseline_data:
        ws.append([
            row['From'], row['To'], 
            round(row['Dist_eGNSS'],3), round(row['Dist_TWD97'],3), round(row['dDist'],4), 
            row['PPM_Text'], row['Check_Dist'],
            round(row['Az_eGNSS'],4), round(row['Az_TWD97'],4), round(row['dAzi_Sec'],1), row['Check_Azi']
        ])
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

def generate_report_7_field(check_results):
    wb = Workbook(); ws = wb.active; ws.title = "實地檢測報表"
    ws.merge_cells('A1:G1'); ws['A1'] = "實地檢測成果比較表 (距離檢核)"; ws['A1'].font = Font(size=16, bold=True); ws['A1'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A2:G2'); ws['A2'] = "合格標準：水平距離較差 < 0.03m 或 相對誤差 < 1/3000"
    ws['A2'].font = Font(color="0000FF", bold=True); ws['A2'].alignment = Alignment(horizontal='left')
    headers = ["起點", "終點", "實測距離(m)", "成果反算距離(m)", "距離較差(m)", "相對誤差", "判定"]
    setup_excel_style(ws, headers, row_idx=3)
    fail_font = Font(color="FF0000", bold=True); pass_font = Font(color="008000", bold=True)
    for i, row in enumerate(check_results, start=4):
        ws.append([row['From'], row['To'], round(row['Dist_Meas'],3), round(row['Dist_Calc'],3), round(row['dDist'],4), row['Rel_Error'], row['Status']])
        ws.cell(row=i, column=7).font = fail_font if row['Status']=="不合格" else pass_font
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

# 🔥 計算模組
def calculate_residual_correction(tn, te, df_res, power=2):
    num_n, num_e, den = 0, 0, 0
    for _, row in df_res.iterrows():
        dist = math.sqrt((tn - row['N_轉換(GPS)'])**2 + (te - row['E_轉換(GPS)'])**2)
        if dist < 0.001: return row['VX'], row['VY']
        w = 1 / (dist ** power)
        num_n += w * row['VX']; num_e += w * row['VY']; den += w
    if den == 0: return 0, 0
    return num_n / den, num_e / den

def compute_6_parameters(obs, known):
    df_o = pd.DataFrame(obs).set_index('測點名稱'); df_k = pd.DataFrame(known).set_index('測點名稱')
    com = df_o.join(df_k, lsuffix='_obs', rsuffix='_known', how='inner')
    n = len(com)
    if n < 3: return None, None, None, None, n
    A = np.zeros((2*n, 6)); L = np.zeros((2*n, 1))
    for i in range(n):
        r = com.iloc[i]
        A[2*i,:] = [r['N_obs'], r['E_obs'], 1, 0, 0, 0]; L[2*i,0] = r['N_known']
        A[2*i+1,:] = [0, 0, 0, r['N_obs'], r['E_obs'], 1]; L[2*i+1,0] = r['E_known']
    try:
        X, _, _, _ = np.linalg.lstsq(A, L, rcond=None)
        p = X.flatten(); V = A @ X - L
        rmse = np.sqrt(np.sum(V**2)/(2*n-6)) if (2*n-6)>0 else 0
        res = []
        for i in range(n):
            r = com.iloc[i]
            nt = p[0]*r['N_obs'] + p[1]*r['E_obs'] + p[2]; et = p[3]*r['N_obs'] + p[4]*r['E_obs'] + p[5]
            vx = nt - r['N_known']; vy = et - r['E_known']
            res.append({'測點名稱': com.index[i], 'N_已知(Ground)': r['N_known'], 'E_已知(Ground)': r['E_known'], 'N_轉換(GPS)': nt, 'E_轉換(GPS)': et, 'VX': vx, 'VY': vy, '平面殘差': np.sqrt(vx**2+vy**2)})
        return p, pd.DataFrame(res), rmse, np.sum(V**2), n
    except Exception as e: return None, None, None, None, str(e)

def generate_report_5_mimic(params, df_res, rmse, sum_vv): 
    wb = Workbook(); ws = wb.active; ws.title = "轉換參數報表"
    align_c = Alignment(horizontal='center', vertical='center'); bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    curr = 1
    ws.merge_cells(f'A{curr}:C{curr}'); ws[f'A{curr}']="RESIDUALS TABLE (單位: m)"; ws[f'A{curr}'].font=bold_font; ws[f'A{curr}'].alignment=align_c; curr+=1
    headers_1 = ["NAME", "VX", "VY"]
    for i, h in enumerate(headers_1, 1):
        cell = ws.cell(row=curr, column=i); cell.value=h; cell.border=border; cell.alignment=align_c; cell.font=bold_font; cell.fill=PatternFill(start_color="D3D3D3", fill_type="solid")
    curr+=1
    for _, r in df_res.iterrows():
        ws.cell(row=curr, column=1, value=r['測點名稱']).border=border; ws.cell(row=curr, column=2, value=round(r['VX'], 6)).border=border; ws.cell(row=curr, column=3, value=round(r['VY'], 6)).border=border; curr+=1
    curr+=2
    ws[f'A{curr}'] = f"SUM OF [VV] = {sum_vv:.6f}"; curr+=1 
    ws[f'A{curr}'] = f"DEGREE OF FREEDOM = {2*len(df_res)-6}"; curr+=1
    ws[f'A{curr}'] = f"STANDARD ERROR = {rmse:.6f} [M]"; curr+=2
    ws[f'A{curr}'] = "TRANSFORMATION PARAMETERS (6-Param Affine)"; curr+=1
    p_names = ["a", "b", "c", "d", "e", "f"]
    for i, p_val in enumerate(params): ws[f'A{curr}'] = f"{p_names[i]} ="; ws[f'B{curr}'] = f"{p_val:.10f}"; curr+=1
    curr+=2
    ws.merge_cells(f'A{curr}:G{curr}'); ws[f'A{curr}']="DISTANCE CHECK (邊長檢核)"; ws[f'A{curr}'].font=bold_font; ws[f'A{curr}'].alignment=align_c; curr+=1
    headers_d = ["FROM", "TO", "GPS", "GROUND", "DIFFERENCE", "1/ PPM", "TEST"]
    for i, h in enumerate(headers_d, 1):
        cell=ws.cell(row=curr, column=i); cell.value=h; cell.border=border; cell.alignment=align_c; cell.fill=PatternFill(start_color="D3D3D3", fill_type="solid"); cell.font=bold_font
    curr+=1
    points = df_res.to_dict('records'); pairs = list(itertools.combinations(points, 2))
    for p1, p2 in pairs:
        dg = math.sqrt((p1['N_轉換(GPS)']-p2['N_轉換(GPS)'])**2+(p1['E_轉換(GPS)']-p2['E_轉換(GPS)'])**2)
        dk = math.sqrt((p1['N_已知(Ground)']-p2['N_已知(Ground)'])**2+(p1['E_已知(Ground)']-p2['E_已知(Ground)'])**2)
        diff = dg - dk; ppm = int(dk/abs(diff)) if abs(diff)>0 else 0
        ws.cell(row=curr,column=1,value=p1['測點名稱']).border=border; ws.cell(row=curr,column=2,value=f"---> {p2['測點名稱']}").border=border; ws.cell(row=curr,column=3,value=round(dg,4)).border=border
        ws.cell(row=curr,column=4,value=round(dk,4)).border=border; ws.cell(row=curr,column=5,value=round(diff,4)).border=border; ws.cell(row=curr,column=6,value=f"{ppm}" if diff!=0 else "Inf").border=border
        ws.cell(row=curr,column=7,value="OK" if (diff==0 or ppm>5000) else "Check").border=border; curr+=1
    curr+=2
    ws.merge_cells(f'A{curr}:F{curr}'); ws[f'A{curr}']="AZIMUTH CHECK (方位角檢核)"; ws[f'A{curr}'].font=bold_font; ws[f'A{curr}'].alignment=align_c; curr+=1
    headers_a = ["FROM", "TO", "GPS", "GROUND", "DIFFERENCE(SEC)", "TEST"]
    for i, h in enumerate(headers_a, 1):
        cell=ws.cell(row=curr, column=i); cell.value=h; cell.border=border; cell.alignment=align_c; cell.fill=PatternFill(start_color="D3D3D3", fill_type="solid"); cell.font=bold_font
    curr+=1
    for p1, p2 in pairs:
        azg = (math.degrees(math.atan2(p2['E_轉換(GPS)']-p1['E_轉換(GPS)'], p2['N_轉換(GPS)']-p1['N_轉換(GPS)']))+360)%360
        azk = (math.degrees(math.atan2(p2['E_已知(Ground)']-p1['E_已知(Ground)'], p2['N_已知(Ground)']-p1['N_已知(Ground)']))+360)%360
        diff_sec = (azg - azk) * 3600
        if diff_sec > 180*3600: diff_sec -= 360*3600
        if diff_sec < -180*3600: diff_sec += 360*3600
        ws.cell(row=curr,column=1,value=p1['測點名稱']).border=border; ws.cell(row=curr,column=2,value=f"---> {p2['測點名稱']}").border=border
        ws.cell(row=curr,column=3,value=deg_to_dmmss(azg)).number_format='0.000000'; ws.cell(row=curr,column=3).border=border
        ws.cell(row=curr,column=4,value=deg_to_dmmss(azk)).number_format='0.000000'; ws.cell(row=curr,column=4).border=border
        ws.cell(row=curr,column=5,value=round(diff_sec,2)).border=border; ws.cell(row=curr,column=6,value="OK" if abs(diff_sec)<20 else "Check").border=border; curr+=1
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

def generate_report_known_check_twd97(df_res):
    wb = Workbook(); ws = wb.active; ws.title = "已知點檢測"
    align_c = Alignment(horizontal='center', vertical='center'); bold_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ws.merge_cells('A1:H1'); ws['A1']="內政部土地測量局已知點檢測成果報表"; ws['A1'].font=Font(size=14, bold=True); ws['A1'].alignment=align_c
    ws.merge_cells('A2:A3'); ws['A2']="點號"; ws.merge_cells('B2:C2'); ws['B2']="自由網坐標"; ws['B3']="N-坐標(m)"; ws['C3']="E-坐標(m)"
    ws.merge_cells('D2:E2'); ws['D2']="已知點坐標"; ws['D3']="N-坐標(m)"; ws['E3']="E-坐標(m)"; ws.merge_cells('F2:H2'); ws['F2']="較差"; ws['F3']="dN(m)"; ws['G3']="dE(m)"; ws['H3']="差值"
    for r in [2, 3]:
        for c in range(1, 9): cell = ws.cell(row=r, column=c); cell.border=border; cell.alignment=align_c; cell.font=bold_font
    for i, row in enumerate(df_res.to_dict('records'), 4):
        dn = row['VX']; de = row['VY']; diff_val = math.sqrt(dn**2 + de**2)
        ws.cell(row=i, column=1, value=row['測點名稱']).border=border; ws.cell(row=i, column=2, value=round(row['N_轉換(GPS)'], 3)).border=border
        ws.cell(row=i, column=3, value=round(row['E_轉換(GPS)'], 3)).border=border; ws.cell(row=i, column=4, value=round(row['N_已知(Ground)'], 3)).border=border
        ws.cell(row=i, column=5, value=round(row['E_已知(Ground)'], 3)).border=border; ws.cell(row=i, column=6, value=round(dn, 3)).border=border
        ws.cell(row=i, column=7, value=round(de, 3)).border=border; ws.cell(row=i, column=8, value=round(diff_val, 3)).border=border
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

def generate_residuals_report(df):
    wb = Workbook(); ws = wb.active; ws.title="殘差表"
    setup_excel_style(ws, ["測點名稱", "N已知", "E已知", "N轉換", "E轉換", "VX", "VY", "平面殘差"], row_idx=1)
    for _, r in df.iterrows(): ws.append([r['測點名稱'], r['N_已知(Ground)'], r['E_已知(Ground)'], r['N_轉換(GPS)'], r['E_轉換(GPS)'], r['VX'], r['VY'], r['平面殘差']])
    adjust_col_width(ws); f = io.BytesIO(); wb.save(f); f.seek(0); return f

# ================= UI 邏輯 =================
with col_right:
    st.header("📝 第四區：檢核流程與暫存")
    c1, c2 = st.columns(2)
    if c1.button("1️⃣ 檢核 第 1 測回", use_container_width=True): 
        if files_round1:
            d, l = process_single_round(files_round1, "第 1 測回")
            if d is not None: 
                st.session_state.current_stage_data = d
                st.session_state.logs = l
                st.session_state.current_stage_name = "第 1 測回"
                st.session_state.show_overwrite_warning = False
            else:
                for err in l: st.error(err)
        else: st.error("請先上傳第 1 測回檔案")
        
    if c2.button("2️⃣ 檢核 第 2 測回", use_container_width=True):
        if files_round2:
            d, l = process_single_round(files_round2, "第 2 測回")
            if d is not None: 
                st.session_state.current_stage_data = d
                st.session_state.logs = l
                st.session_state.current_stage_name = "第 2 測回"
                st.session_state.show_overwrite_warning = False
            else:
                for err in l: st.error(err)
        else: st.error("請先上傳第 2 測回檔案")
    
    if st.session_state.current_stage_name != "":
        st.markdown("---")
        st.info(f"預覽: {st.session_state.current_stage_name}")
        with st.expander("檢核日誌", expanded=True):
            for line in st.session_state.logs: 
                if "✅" in line: st.success(line)
                elif "❌" in line: st.error(line)
                else: st.write(line)
                
        if st.session_state.current_stage_data is not None and len(st.session_state.current_stage_data) > 0:
            st.dataframe(pd.DataFrame(st.session_state.current_stage_data).head(3))
            
            if 'show_overwrite_warning' not in st.session_state: st.session_state.show_overwrite_warning = False
            if 'pending_data_to_write' not in st.session_state: st.session_state.pending_data_to_write = []
            if 'duplicate_list' not in st.session_state: st.session_state.duplicate_list = []

            if st.button("💾 寫入暫存庫", type="primary", use_container_width=True):
                dups = [f"{r['測點名稱']} ({r['測回別']})" for r in st.session_state.current_stage_data if any((d['測點名稱']==r['測點名稱'] and d['測回別']==r['測回別']) for d in st.session_state.temp_database)]
                if dups: 
                    st.session_state.show_overwrite_warning=True
                    st.session_state.pending_data_to_write=st.session_state.current_stage_data
                    st.session_state.duplicate_list=dups
                else: 
                    st.session_state.temp_database.extend(st.session_state.current_stage_data)
                    st.success("✅ 寫入成功")
                    st.session_state.current_stage_data=None
                    st.session_state.current_stage_name="" 
                    st.rerun()

            if st.session_state.show_overwrite_warning:
                st.warning("⚠️ 資料重複: " + ", ".join(st.session_state.duplicate_list))
                cc1, cc2 = st.columns(2)
                if cc1.button("✅ 覆蓋"):
                    for r in st.session_state.pending_data_to_write:
                        st.session_state.temp_database = [d for d in st.session_state.temp_database if not (d['測點名稱']==r['測點名稱'] and d['測回別']==r['測回別'])]
                    st.session_state.temp_database.extend(st.session_state.pending_data_to_write)
                    st.success("已覆蓋")
                    st.session_state.current_stage_data=None
                    st.session_state.current_stage_name=""
                    st.session_state.show_overwrite_warning=False
                    st.rerun()
                if cc2.button("❌ 取消"): 
                    st.session_state.show_overwrite_warning=False
                    st.rerun()
        else:
            st.warning("⚠️ 目前沒有任何測點通過檢核，請查看上方日誌了解原因。")

    st.markdown("---"); st.subheader("🗄️ 暫存資料庫")
    if st.session_state.temp_database:
        df_db = pd.DataFrame(st.session_state.temp_database)
        if "移除" not in df_db.columns: df_db.insert(0, "移除", False)
        edited_df = st.data_editor(
            df_db, 
            hide_index=True, 
            use_container_width=True,
            disabled=["測點名稱", "測回別", "有效筆數", "總計點數", "使用比率", "N", "E", "H", "sN", "sE", "sH", "儀器高"]
        )
        cd1, cd2 = st.columns([1,4])
        if cd1.button("🗑️ 刪除"): st.session_state.temp_database = edited_df[edited_df['移除']==False].drop(columns=['移除']).to_dict('records'); st.rerun()
        if cd2.button("💣 清空"): st.session_state.temp_database=[]; st.rerun()
    else: st.info("暫存區為空")

    st.markdown("---"); st.header("📤 第六區：報表產出")
    if st.session_state.temp_database:
        r6_limits = {'diff_inst': diff_inst_limit, 'diff_plane': diff_plane_limit, 'diff_elev': diff_elev_limit, 'time_gap': time_gap_limit}
        cr1, cr2, cr3 = st.columns(3)
        with cr1: 
            f6 = generate_report_6_final(st.session_state.temp_database, r6_limits)
            st.download_button("📊 下載 報表6-1.精度檢核總表", f6, "報表6-1.xlsx", use_container_width=True)
        with cr2:
            f6_2 = generate_report_6_2_center(st.session_state.temp_database)
            st.download_button("📝 下載 報表6-2.測繪中心檢核表", f6_2, "報表6-2_測繪中心版.xlsx", use_container_width=True)
        with cr3:
            f2 = generate_report_2_coord(st.session_state.temp_database)
            st.download_button("📍 下載 報表2.坐標成果表", f2, "報表2.xlsx", use_container_width=True)

    st.markdown("---"); st.header("🏁 第七區：六參數轉換 (強制附合)")
    kp_file = st.file_uploader("📂 上傳 已知控制點清冊", type=['csv', 'xlsx'], key="kp_u")
    if kp_file and st.session_state.temp_database:
        try:
            if kp_file.name.endswith('.csv'): df_kp = pd.read_csv(kp_file, encoding='utf-8-sig')
            else: df_kp = pd.read_excel(kp_file)
            df_kp.columns = df_kp.columns.str.replace('\n', '').str.replace('"', '').str.strip()
            if {'測點名稱','N','E'}.issubset(df_kp.columns):
                k_list = df_kp[['測點名稱','N','E']].to_dict('records')
                obs_avg = pd.DataFrame(st.session_state.temp_database).groupby('測點名稱')[['N','E']].mean().reset_index().to_dict('records')
                if st.button("🚀 解算六參數", type="primary", use_container_width=True):
                    p, df_res, rmse, vv, n = compute_6_parameters(obs_avg, k_list)
                    if p is not None:
                        st.session_state.trans_params=p; st.session_state.trans_residuals=df_res; st.session_state.trans_rmse=rmse; st.session_state.trans_vv=vv
                        st.success(f"✅ 解算完成 (共{n}點)")
                    else: st.error(f"解算失敗: {n}")
            else: st.error("欄位需包含: 測點名稱, N, E")
        except Exception as e: st.error(str(e))

    if st.session_state.trans_params is not None:
        st.markdown("### 📊 轉換評估")
        cols_order = ['測點名稱', 'N_已知(Ground)', 'E_已知(Ground)', 'N_轉換(GPS)', 'E_轉換(GPS)', 'VX', 'VY', '平面殘差']
        df_show = st.session_state.trans_residuals[cols_order].copy()
        
        st.dataframe(df_show.style.format({
            'N_已知(Ground)': '{:.4f}', 'N_轉換(GPS)': '{:.4f}', 'E_已知(Ground)': '{:.4f}', 'E_轉換(GPS)': '{:.4f}',
            'VX': '{:.4f}', 'VY': '{:.4f}', '平面殘差': '{:.4f}'
        }))
        f_res = generate_residuals_report(df_show)
        st.download_button("📥 下載 已知點殘差比較表.xlsx", f_res, "已知點殘差比較表.xlsx", use_container_width=True)
        
        st.metric("轉換 RMSE", f"{st.session_state.trans_rmse:.4f} m", help=f"門檻: {rmse_threshold}m")
        ok_apply = st.session_state.trans_rmse <= rmse_threshold
        if ok_apply: st.success("✅ 精度符合，建議全區轉換")
        else: st.error("⚠️ 精度不足，請檢查")
        
        f5 = generate_report_5_mimic(st.session_state.trans_params, st.session_state.trans_residuals, st.session_state.trans_rmse, st.session_state.trans_vv)
        st.download_button("📊 下載 報表5.參數轉換報表.xlsx", f5, "報表5.xlsx")

        st.markdown("### 🌍 全區應用 (強制附合)")
        if st.button("🚀 執行全區轉換 (TWD97)", type="primary", disabled=not ok_apply, use_container_width=True):
            p = st.session_state.trans_params
            obs_all = pd.DataFrame(st.session_state.temp_database).groupby('測點名稱')[['N','E','H']].mean().reset_index()
            res = []
            for _, r in obs_all.iterrows():
                na = p[0]*r['N'] + p[1]*r['E'] + p[2]; ea = p[3]*r['N'] + p[4]*r['E'] + p[5]
                dn, de = calculate_residual_correction(na, ea, st.session_state.trans_residuals)
                res.append({'測點名稱': r['測點名稱'], 'N_TWD97': na+dn, 'E_TWD97': ea+de, 'H': r['H']})
            st.session_state.final_twd97_data = res
            st.success("✅ 全區轉換完成")

        if st.session_state.final_twd97_data:
            st.dataframe(pd.DataFrame(st.session_state.final_twd97_data).head())
            c_rep1, c_rep2 = st.columns(2)
            with c_rep1:
                f3 = generate_report_3_twd97(st.session_state.final_twd97_data)
                st.download_button("📥 下載 報表3.TWD97成果表.xlsx", f3, "報表3.xlsx", type="primary", use_container_width=True)
            with c_rep2:
                f_check = generate_report_known_check_twd97(st.session_state.trans_residuals)
                st.download_button("📥 下載 已知點檢測成果報表.xlsx", f_check, "已知點檢測成果報表.xlsx", type="primary", use_container_width=True)

    # ================= 8. 基線比較與實地檢測 =================
    if st.session_state.final_twd97_data:
        st.markdown("---")
        st.header("🔍 第八區：基線比較與實地檢測")
        
        if st.button("🚀 計算全組合基線 (e-GNSS vs TWD97)", type="primary", use_container_width=True):
            df_gnss = pd.DataFrame(st.session_state.temp_database).groupby('測點名稱')[['N','E']].mean()
            df_twd97 = pd.DataFrame(st.session_state.final_twd97_data).set_index('測點名稱')[['N_TWD97','E_TWD97']]
            common_pts = sorted(list(set(df_gnss.index) & set(df_twd97.index)))
            
            results = []
            for p1, p2 in itertools.combinations(common_pts, 2):
                d_g, az_g = calc_dist_azimuth(df_gnss.loc[p1,'N'], df_gnss.loc[p1,'E'], df_gnss.loc[p2,'N'], df_gnss.loc[p2,'E'])
                d_t, az_t = calc_dist_azimuth(df_twd97.loc[p1,'N_TWD97'], df_twd97.loc[p1,'E_TWD97'], df_twd97.loc[p2,'N_TWD97'], df_twd97.loc[p2,'E_TWD97'])
                
                d_diff = abs(d_t - d_g)
                inv_p = int(d_t / d_diff) if d_diff > 0.0001 else 99999
                check_d = "合格" if inv_p >= 5000 else "不合格"
                
                az_diff = az_t - az_g
                if az_diff > 180: az_diff -= 360
                if az_diff < -180: az_diff += 360
                az_diff_sec = abs(az_diff * 3600)
                check_a = "合格" if az_diff_sec <= 20 else "不合格"
                
                results.append({
                    'From': p1, 'To': p2,
                    'Dist_eGNSS': d_g, 'Dist_TWD97': d_t, 'dDist': d_diff,
                    'PPM_Text': f"1/{inv_p}" if inv_p < 99999 else "無限大", 'Check_Dist': check_d,
                    'Az_eGNSS': az_g, 'Az_TWD97': az_t, 'dAzi_Sec': az_diff_sec, 'Check_Azi': check_a
                })
            
            st.session_state.baseline_check_data = results
            st.success(f"已完成 {len(results)} 組基線比對！")

        if st.session_state.baseline_check_data:
            df_base = pd.DataFrame(st.session_state.baseline_check_data)
            st.write("▼ **報表 4 預覽 (全組合基線比較):**")
            st.dataframe(df_base.style.format({
                'Dist_eGNSS': '{:.3f}', 'Dist_TWD97': '{:.3f}', 'dDist': '{:.4f}',
                'Az_eGNSS': '{:.4f}', 'Az_TWD97': '{:.4f}', 'dAzi_Sec': '{:.1f}'
            }))
            f4 = generate_report_4_baseline(st.session_state.baseline_check_data)
            st.download_button("📊 下載 報表4.全組合基線比較表", f4, "報表4.xlsx", type="primary", use_container_width=True)
            
            st.markdown("#### 📏 實地檢測建議清單 (短邊 < 100m)")
            short_baselines = df_base[df_base['Dist_TWD97'] <= short_dist_limit].copy()
            
            if not short_baselines.empty:
                st.warning(f"⚠️ 發現 {len(short_baselines)} 組短邊，建議進行實地檢測！")
                st.dataframe(short_baselines[['From', 'To', 'Dist_TWD97']].style.format({'Dist_TWD97': '{:.3f}'}))
                
                sample_data = short_baselines[['From', 'To']].copy()
                sample_data['實測距離(m)'] = ""
                f_sample = io.BytesIO()
                with pd.ExcelWriter(f_sample, engine='openpyxl') as writer:
                    sample_data.to_excel(writer, index=False)
                f_sample.seek(0)
                st.download_button("📥 下載 外業檢測紀錄表範例檔.xlsx", f_sample, "外業檢測範例.xlsx")
                
                st.markdown("#### 📤 上傳實測成果進行比對")
                field_file = st.file_uploader("上傳填寫完成的檢測表", type=['xlsx'])
                
                if field_file:
                    try:
                        df_field = pd.read_excel(field_file)
                        check_res = []
                        for _, row in df_field.iterrows():
                            p1, p2 = str(row['From']), str(row['To'])
                            match = short_baselines[(short_baselines['From']==p1) & (short_baselines['To']==p2)]
                            if not match.empty:
                                d_calc = match.iloc[0]['Dist_TWD97']
                                d_meas = row.get('實測距離(m)', 0)
                                if pd.notnull(d_meas):
                                    dd = abs(d_meas - d_calc)
                                    is_pass = False
                                    if dd <= 0.03: is_pass = True
                                    elif d_calc > 0 and (d_calc / dd) >= 3000: is_pass = True
                                    
                                    status = "合格" if is_pass else "不合格"
                                    rel = f"1/{int(d_calc/dd)}" if dd > 0.001 else "無限大"
                                    
                                    check_res.append({
                                        'From': p1, 'To': p2,
                                        'Dist_Meas': d_meas, 'Dist_Calc': d_calc, 'dDist': dd, 'Rel_Error': rel,
                                        'Status': status
                                    })
                        
                        if check_res:
                            st.write("▼ **實地檢測成果比較:**")
                            st.dataframe(pd.DataFrame(check_res))
                            f7 = generate_report_7_field(check_res)
                            st.download_button("📊 下載 實地檢測報表", f7, "實地檢測報表.xlsx", type="primary")
                            
                    except Exception as e: st.error(f"讀取錯誤: {e}")
            else:
                st.success("✅ 無小於 100m 之短邊，無需強制實地檢測。")
