import streamlit as st
import sqlite3
import pandas as pd
import openpyxl
from datetime import datetime
import os

st.set_page_config(
    page_title="MRO Stock Management System | TPS Standard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Professional Industrial Theme (Toyota Production System inspired)
st.markdown('''
<style>
    /* CSS Reset & Variables */
    :root {
        --tps-red: #D1001C;
        --tps-dark-slate: #1F2937;
        --tps-slate-bg: #F3F4F6;
        --tps-card-bg: #FFFFFF;
        --tps-green: #10B981;
        --tps-blue: #2563EB;
        --tps-amber: #F59E0B;
        --tps-border: #E5E7EB;
    }
    
    .stApp {
        background-color: #F8FAFC;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }

    /* Top Bar Header Style */
    .tps-header {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        padding: 20px 28px;
        border-radius: 12px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        border-left: 6px solid var(--tps-red);
    }
    .tps-header h1 {
        color: #FFFFFF !important;
        font-size: 24px !important;
        font-weight: 800 !important;
        letter-spacing: 0.5px;
        margin: 0 !important;
        padding: 0 !important;
    }
    .tps-header p {
        color: #94A3B8 !important;
        margin: 4px 0 0 0 !important;
        font-size: 13px !important;
        font-weight: 500;
    }

    /* Metric Cards (Andon Board concept) */
    .metric-card {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 16px 20px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: transform 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
    }
    .metric-title {
        font-size: 12px;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 24px;
        font-weight: 800;
        color: #0F172A;
        margin-top: 4px;
    }

    /* Form Container (5S / Visual Management) */
    .tps-form-container {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 24px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        margin-bottom: 20px;
    }

    /* Custom Streamlit Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #E2E8F0;
        padding: 6px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        border-radius: 8px;
        font-weight: 700;
        font-size: 14px;
        color: #475569;
        background-color: transparent;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
</style>
''', unsafe_allow_html=True)

EXCEL_DANH_MUC = "danh_muc_mro.xlsx"
DB_FILE = "mro_production.db"

# ==========================================
# 1. DATABASE & SETUP
# ==========================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS inventory (item_num TEXT PRIMARY KEY, ma_hang TEXT, ten_eng TEXT, ten_vie TEXT, ton_kho INTEGER DEFAULT 0, min_safety INTEGER DEFAULT 5)")
    cursor.execute("CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY AUTOINCREMENT, loai_gd TEXT, item_num TEXT, ma_hang TEXT, ten_vie TEXT, so_luong INTEGER, nguoi_thuc_hien TEXT, ghi_chu TEXT, ngay_gio TEXT)")
    conn.commit()
    conn.close()

def load_and_sync_danh_muc():
    init_db()
    if os.path.exists(EXCEL_DANH_MUC):
        try:
            df = pd.read_excel(EXCEL_DANH_MUC)
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            for _, row in df.iterrows():
                item_num = str(row.iloc[0]).strip().upper() if pd.notna(row.iloc[0]) else ""
                ma_hang = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                ten_eng = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                ten_vie = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                
                if item_num:
                    cursor.execute("INSERT INTO inventory (item_num, ma_hang, ten_eng, ten_vie, ton_kho) VALUES (?, ?, ?, ?, 0) ON CONFLICT(item_num) DO UPDATE SET ma_hang=excluded.ma_hang, ten_eng=excluded.ten_eng, ten_vie=excluded.ten_vie", (item_num, ma_hang, ten_eng, ten_vie))
            conn.commit()
            conn.close()
        except Exception as e:
            st.error(f"Lỗi đồng bộ danh mục Excel: {e}")

load_and_sync_danh_muc()

def get_connection():
    return sqlite3.connect(DB_FILE)

# ==========================================
# 2. HEADER & VISUAL MANAGEMENT (ANDON BOARD)
# ==========================================
st.markdown('''
<div class="tps-header">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h1>⚙️ HỆ THỐNG QUẢN LÝ KHO MRO - FORMING</h1>
            <p>QUẢN LÝ TRỰC QUAN (VISUAL MANAGEMENT) | BẢNG KANBAN KHO PRODUCTION</p>
        </div>
        <div style="text-align: right;">
            <span style="background: #2563EB; color: white; padding: 4px 12px; border-radius: 20px; font-size: 11px; font-weight: 700; letter-spacing: 0.5px;">SYSTEM LIVE</span>
        </div>
    </div>
</div>
''', unsafe_allow_html=True)

# Calculation for Quick KPIs / Andon Summary
conn = get_connection()
df_inv_summary = pd.read_sql_query("SELECT item_num, ton_kho, min_safety FROM inventory", conn)
df_tx_today = pd.read_sql_query("SELECT loai_gd, so_luong FROM transactions WHERE DATE(ngay_gio) = DATE('now', 'localtime')", conn)
conn.close()

total_items = len(df_inv_summary)
low_stock_items = len(df_inv_summary[df_inv_summary['ton_kho'] <= df_inv_summary['min_safety']]) if not df_inv_summary.empty else 0
today_inbound = df_tx_today[df_tx_today['loai_gd'] == 'NHẬP']['so_luong'].sum() if not df_tx_today.empty else 0
today_outbound = df_tx_today[df_tx_today['loai_gd'] == 'XUẤT']['so_luong'].sum() if not df_tx_today.empty else 0

# Quick Metrics KPI Row (Andon Indicator)
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
with kpi1:
    st.markdown(f'''
    <div class="metric-card" style="border-left: 4px solid #2563EB;">
        <div class="metric-title">Tổng Danh Mục MRO</div>
        <div class="metric-value">{total_items} <span style="font-size:13px; font-weight:normal; color:#64748B;">SKU</span></div>
    </div>
    ''', unsafe_allow_html=True)

with kpi2:
    alert_color = "#DC2626" if low_stock_items > 0 else "#10B981"
    st.markdown(f'''
    <div class="metric-card" style="border-left: 4px solid {alert_color};">
        <div class="metric-title">Cảnh Báo Tồn Thấp (Kanban)</div>
        <div class="metric-value" style="color: {alert_color};">{low_stock_items} <span style="font-size:13px; font-weight:normal; color:#64748B;">Item</span></div>
    </div>
    ''', unsafe_allow_html=True)

with kpi3:
    st.markdown(f'''
    <div class="metric-card" style="border-left: 4px solid #10B981;">
        <div class="metric-title">Nhập Trong Ngày</div>
        <div class="metric-value" style="color: #059669;">+{today_inbound} <span style="font-size:13px; font-weight:normal; color:#64748B;">Pcs</span></div>
    </div>
    ''', unsafe_allow_html=True)

with kpi4:
    st.markdown(f'''
    <div class="metric-card" style="border-left: 4px solid #F59E0B;">
        <div class="metric-title">Xuất Trong Ngày</div>
        <div class="metric-value" style="color: #D97706;">-{today_outbound} <span style="font-size:13px; font-weight:normal; color:#64748B;">Pcs</span></div>
    </div>
    ''', unsafe_allow_html=True)

st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)

# ==========================================
# 3. MAIN WORKFLOW TABS
# ==========================================
tab_kanban, tab_nhap, tab_xuat, tab_history = st.tabs([
    "📊 TỒN KHO REALTIME & KANBAN", 
    "📥 NHẬP KHO (INBOUND)", 
    "📤 XUẤT KHO (OUTBOUND)", 
    "📜 LỊCH SỬ GIAO DỊCH"
])

# ------------------------------------------
# TAB 1: KANBAN & REALTIME INVENTORY
# ------------------------------------------
with tab_kanban:
    st.markdown("### 🔍 BẢNG ĐIỀU HÀNH TỒN KHO (VISUAL KANBAN BOARD)")
    
    conn = get_connection()
    df_inventory = pd.read_sql_query("SELECT item_num AS 'Item#', ma_hang AS 'Mã Hàng', ten_vie AS 'Tên Material', ten_eng AS 'Tên Tiếng Anh', ton_kho AS 'Tồn Kho Realtime', min_safety AS 'Mức An Toàn' FROM inventory ORDER BY ton_kho ASC", conn)
    conn.close()

    # Search bar & Filters
    col_s1, col_s2 = st.columns([3, 1.5])
    with col_s1:
        search_kw = st.text_input("🔍 Tìm kiếm theo Item#, Mã Hàng, Tên Material...", key="inv_search").lower().strip()
    with col_s2:
        stock_filter = st.selectbox("Trạng thái Tồn kho", ["Tất cả", "⚠️ Cảnh báo thấp (Kanban)", "✅ Đủ tồn kho"])

    # Filter logic
    df_filtered = df_inventory.copy()
    if search_kw and not df_filtered.empty:
        mask = df_filtered.apply(lambda r: r.astype(str).str.lower().str.contains(search_kw).any(), axis=1)
        df_filtered = df_filtered[mask]
    
    if stock_filter == "⚠️ Cảnh báo thấp (Kanban)" and not df_filtered.empty:
        df_filtered = df_filtered[df_filtered["Tồn Kho Realtime"] <= df_filtered["Mức An Toàn"]]
    elif stock_filter == "✅ Đủ tồn kho" and not df_filtered.empty:
        df_filtered = df_filtered[df_filtered["Tồn Kho Realtime"] > df_filtered["Mức An Toàn"]]

    max_val = int(df_inventory["Tồn Kho Realtime"].max()) if not df_inventory.empty and pd.notna(df_inventory["Tồn Kho Realtime"].max()) else 50

    st.dataframe(
        df_filtered,
        column_config={
            "Item#": st.column_config.TextColumn("Item#", width="medium"),
            "Tồn Kho Realtime": st.column_config.ProgressColumn(
                "Mức Tồn Kho Visual",
                format="%d",
                min_value=0,
                max_value=max_val + 10,
            ),
        },
        use_container_width=True,
        hide_index=True,
        height=450
    )

    # Download Button
    csv_data = df_filtered.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(
        label="📗 Xuất Báo Cáo Tồn Kho (Excel/CSV)",
        data=csv_data,
        file_name=f"BAO_CAO_MRO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv"
    )

# ------------------------------------------
# TAB 2: NHẬP KHO (INBOUND)
# ------------------------------------------
with tab_nhap:
    st.markdown("### 📥 THÔNG TIN PHIẾU NHẬP KHO MRO")
    
    with st.container():
        st.markdown('<div class="tps-form-container">', unsafe_allow_html=True)
        with st.form("form_nhap_tps", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                item_nhap = st.text_input("Mã Item# (*)").strip().upper()
                sl_nhap = st.number_input("Số lượng nhập (*)", min_value=1, step=1, value=1)
            with c2:
                nguoi_nhap = st.text_input("Người thực hiện / Nhân viên Kho")
                ghi_chu_nhap = st.text_input("Vị trí lưu kho (Location) / Ghi chú")

            st.markdown("<br>", unsafe_allow_html=True)
            submit_nhap = st.form_submit_button("💾 XÁC NHẬN NHẬP KHO", use_container_width=True)

        if submit_nhap:
            if not item_nhap:
                st.error("⚠️ Vui lòng nhập Mã Item#!")
            else:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT ma_hang, ten_vie FROM inventory WHERE item_num = ?", (item_nhap,))
                row = cursor.fetchone()

                if not row:
                    st.error(f"❌ Không tìm thấy Item# [{item_nhap}] trong danh mục MRO!")
                    conn.close()
                else:
                    ma_hang, ten_vie = row
                    ngay_gio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    cursor.execute("UPDATE inventory SET ton_kho = ton_kho + ? WHERE item_num = ?", (sl_nhap, item_nhap))
                    cursor.execute("INSERT INTO transactions (loai_gd, item_num, ma_hang, ten_vie, so_luong, nguoi_thuc_hien, ghi_chu, ngay_gio) VALUES ('NHẬP', ?, ?, ?, ?, ?, ?, ?)", (item_nhap, ma_hang, ten_vie, sl_nhap, nguoi_nhap, ghi_chu_nhap, ngay_gio))
                    conn.commit()
                    conn.close()
                    st.success(f"✅ [SUCCESS] Đã NHẬP **{sl_nhap}** x **{ten_vie}** ({item_nhap}) vào hệ thống!")
                    st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------
# TAB 3: XUẤT KHO (OUTBOUND - PULL SYSTEM)
# ------------------------------------------
with tab_xuat:
    st.markdown("### 📤 THÔNG TIN PHIẾU XUẤT KHO (PULL SYSTEM)")
    
    with st.container():
        st.markdown('<div class="tps-form-container">', unsafe_allow_html=True)
        with st.form("form_xuat_tps", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                item_xuat = st.text_input("Mã Item# (*)").strip().upper()
                sl_xuat = st.number_input("Số lượng xuất (*)", min_value=1, step=1, value=1)
            with c2:
                nguoi_xuat = st.text_input("Người nhận / Bộ phận sản xuất")
                ghi_chu_xuat = st.text_input("Mục đích / Mã máy / Line sản xuất")

            st.markdown("<br>", unsafe_allow_html=True)
            submit_xuat = st.form_submit_button("📤 XÁC NHẬN XUẤT KHO", use_container_width=True)

        if submit_xuat:
            if not item_xuat:
                st.error("⚠️ Vui lòng nhập Mã Item#!")
            else:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT ma_hang, ten_vie, ton_kho FROM inventory WHERE item_num = ?", (item_xuat,))
                row = cursor.fetchone()

                if not row:
                    st.error(f"❌ Không tìm thấy Item# [{item_xuat}] trong kho!")
                    conn.close()
                else:
                    ma_hang, ten_vie, ton_kho_hien_tai = row
                    if sl_xuat > ton_kho_hien_tai:
                        st.error(f"❌ KHO KHÔNG ĐỦ HÀNG! Tồn kho hiện tại chỉ còn [{ton_kho_hien_tai}], không đủ để xuất [{sl_xuat}].")
                        conn.close()
                    else:
                        ngay_gio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        cursor.execute("UPDATE inventory SET ton_kho = ton_kho - ? WHERE item_num = ?", (sl_xuat, item_xuat))
                        cursor.execute("INSERT INTO transactions (loai_gd, item_num, ma_hang, ten_vie, so_luong, nguoi_thuc_hien, ghi_chu, ngay_gio) VALUES ('XUẤT', ?, ?, ?, ?, ?, ?, ?)", (item_xuat, ma_hang, ten_vie, sl_xuat, nguoi_xuat, ghi_chu_xuat, ngay_gio))
                        conn.commit()
                        conn.close()
                        st.success(f"✅ [SUCCESS] Đã XUẤT KHO **{sl_xuat}** x **{ten_vie}** thành công!")
                        st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------
# TAB 4: LỊCH SỬ GIAO DỊCH
# ------------------------------------------
with tab_history:
    st.markdown("### 📜 LỊCH SỬ GIAO DỊCH (TRACEABILITY LOG)")
    
    conn = get_connection()
    df_logs = pd.read_sql_query("SELECT loai_gd AS 'Loại GD', item_num AS 'Item#', ma_hang AS 'Mã Hàng', ten_vie AS 'Tên Material', so_luong AS 'Số Lượng', nguoi_thuc_hien AS 'Người Thực Hiện', ghi_chu AS 'Ghi Chú', ngay_gio AS 'Ngày Giờ' FROM transactions ORDER BY id DESC", conn)
    conn.close()

    # Search & Filter controls
    f_col1, f_col2 = st.columns([1, 3])
    with f_col1:
        type_filter = st.selectbox("Lọc Loại GD", ["Tất cả", "NHẬP", "XUẤT"], key="log_type_filter")
    with f_col2:
        kw_filter = st.text_input("🔍 Từ khóa tìm kiếm (Người thực hiện, Ghi chú, Item#...)", key="log_kw_filter").lower().strip()

    df_logs_filtered = df_logs.copy()
    if type_filter != "Tất cả" and not df_logs_filtered.empty:
        df_logs_filtered = df_logs_filtered[df_logs_filtered["Loại GD"] == type_filter]
    
    if kw_filter and not df_logs_filtered.empty:
        mask = df_logs_filtered.apply(lambda r: r.astype(str).str.lower().str.contains(kw_filter).any(), axis=1)
        df_logs_filtered = df_logs_filtered[mask]

    st.dataframe(
        df_logs_filtered,
        use_container_width=True,
        hide_index=True,
        height=400
    )

    csv_logs = df_logs_filtered.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(
        label="📗 Xuất Lịch Sử Giao Dịch (CSV)",
        data=csv_logs,
        file_name=f"LICH_SU_MRO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv"
    )
    import streamlit as st
import pandas as pd

# Giả sử df là DataFrame dữ liệu bạn muốn kết nối với Power Query
# (Ví dụ: df = pd.read_sql_query("SELECT * FROM inventory", conn))

# Lấy tham số từ URL
query_params = st.query_params

# Nếu đường dẫn có chứa parameter ?export=csv
if query_params.get("export") == "csv":
    # Xuất dữ liệu dưới dạng CSV và dừng giao diện Streamlit
    st.write(df.to_csv(index=False))
    st.stop()