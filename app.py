import io
import sqlite3
import pandas as pd
import streamlit as st

# =========================================================
# 1. CẤU HÌNH TRANG WEB
# =========================================================
st.set_page_config(
    page_title="Hệ Thống Quản Lý Kho MRO - Forming",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# 2. KHỞI TẠO VÀ XỬ LÝ CƠ SỞ DỮ LIỆU (SQLITE)
# =========================================================
DB_FILE = "inventory.db"

def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    # Tạo bảng inventory nếu chưa tồn tại
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_num TEXT UNIQUE NOT NULL,
            item_name TEXT NOT NULL,
            ton_kho INTEGER DEFAULT 0,
            min_safety INTEGER DEFAULT 0,
            location TEXT DEFAULT ''
        )
    """)
    
    # Kiểm tra và tự động thêm cột min_safety nếu CSDL cũ chưa có
    cursor.execute("PRAGMA table_info(inventory)")
    columns = [col[1] for col in cursor.fetchall()]
    if "min_safety" not in columns:
        cursor.execute("ALTER TABLE inventory ADD COLUMN min_safety INTEGER DEFAULT 0")
    if "item_name" not in columns:
        cursor.execute("ALTER TABLE inventory ADD COLUMN item_name TEXT DEFAULT ''")
        
    conn.commit()
    conn.close()

# Chạy khởi tạo CSDL
init_db()

# =========================================================
# 3. KẾT NỐI POWER QUERY DÙNG CHO EXCEL (Xử lý ngầm)
# (Link sử dụng trong Power Query: https://your-app.streamlit.app/?export=csv)
# =========================================================
query_params = st.query_params

if query_params.get("export") == "csv":
    conn = get_connection()
    df_export = pd.read_sql_query("SELECT item_num, item_name, ton_kho, min_safety FROM inventory", conn)
    conn.close()
    
    # Xuất thuần text CSV cho Excel đọc trực tiếp
    st.text(df_export.to_csv(index=False))
    st.stop()  # Dừng tại đây, không tải giao diện web

# =========================================================
# 4. TRUY VẤN DỮ LIỆU TỔNG QUAN
# =========================================================
conn = get_connection()
df_inventory = pd.read_sql_query("SELECT * FROM inventory", conn)
conn.close()

# Tính toán các chỉ số KPI
total_sku = len(df_inventory)
if not df_inventory.empty and 'min_safety' in df_inventory.columns:
    low_stock = len(df_inventory[df_inventory['ton_kho'] <= df_inventory['min_safety']])
else:
    low_stock = 0

# =========================================================
# 5. GIAO DIỆN CHÍNH (HEADER & KPI BANNER)
# =========================================================
# Banner Tiêu đề gốc
st.markdown("""
    <div style="background-color: #0d1117; padding: 18px 25px; border-radius: 8px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h2 style="margin:0; font-size: 22px; color: #ffffff; font-weight: 700;">⚙️ HỆ THỐNG QUẢN LÝ KHO MRO - FORMING</h2>
            <p style="margin:4px 0 0 0; color: #8b949e; font-size: 12px; font-weight: 500;">
                QUẢN LÝ TRỰC QUAN (VISUAL MANAGEMENT) | BẢNG KANBAN KHO PRODUCTION
            </p>
        </div>
        <div>
            <span style="background-color: #1f6beb; color: white; padding: 5px 12px; border-radius: 20px; font-size: 11px; font-weight: bold;">SYSTEM LIVE</span>
        </div>
    </div>
""", unsafe_allow_html=True)

# 4 Thẻ KPI Thống kê
col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)

with col_kpi1:
    st.markdown(f"""
        <div style="border-left: 4px solid #1f6beb; background: #ffffff; padding: 12px 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <div style="font-size: 11px; color: #57606a; font-weight: bold; text-transform: uppercase;">TỔNG DANH MỤC MRO</div>
            <div style="font-size: 22px; font-weight: bold; color: #000; margin-top: 4px;">{total_sku} <span style="font-size: 13px; color: #57606a; font-weight: normal;">SKU</span></div>
        </div>
    """, unsafe_allow_html=True)

with col_kpi2:
    st.markdown(f"""
        <div style="border-left: 4px solid #cf222e; background: #ffffff; padding: 12px 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <div style="font-size: 11px; color: #57606a; font-weight: bold; text-transform: uppercase;">CẢNH BÁO TỒN THẤP (KANBAN)</div>
            <div style="font-size: 22px; font-weight: bold; color: #cf222e; margin-top: 4px;">{low_stock} <span style="font-size: 13px; color: #57606a; font-weight: normal;">Item</span></div>
        </div>
    """, unsafe_allow_html=True)

with col_kpi3:
    st.markdown("""
        <div style="border-left: 4px solid #1a7f37; background: #ffffff; padding: 12px 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <div style="font-size: 11px; color: #57606a; font-weight: bold; text-transform: uppercase;">NHẬP TRONG NGÀY</div>
            <div style="font-size: 22px; font-weight: bold; color: #1a7f37; margin-top: 4px;">+0 <span style="font-size: 13px; color: #57606a; font-weight: normal;">Pcs</span></div>
        </div>
    """, unsafe_allow_html=True)

with col_kpi4:
    st.markdown("""
        <div style="border-left: 4px solid #d97706; background: #ffffff; padding: 12px 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <div style="font-size: 11px; color: #57606a; font-weight: bold; text-transform: uppercase;">XUẤT TRONG NGÀY</div>
            <div style="font-size: 22px; font-weight: bold; color: #d97706; margin-top: 4px;">-0 <span style="font-size: 13px; color: #57606a; font-weight: normal;">Pcs</span></div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =========================================================
# 6. ĐIỀU HƯỚNG TABS
# =========================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 TỒN KHO REALTIME & KANBAN", 
    "📥 NHẬP KHO (INBOUND)", 
    "📤 XUẤT KHO (OUTBOUND)", 
    "📜 LỊCH SỬ GIAO DỊCH"
])

# ---------------------------------------------------------
# TAB 1: TỒN KHO REALTIME & KANBAN
# ---------------------------------------------------------
with tab1:
    st.subheader("📋 Bảng Tổng Hợp Tồn Kho")
    st.dataframe(df_inventory, use_container_width=True)
    
    st.markdown("---")
    st.subheader("🔗 Kết nối & Tải dữ liệu")
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_inventory.to_excel(writer, index=False, sheet_name='Inventory')
        excel_data = output.getvalue()

        st.download_button(
            label="📥 Tải Báo Cáo Excel (.xlsx)",
            data=excel_data,
            file_name="mro_inventory_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with col_d2:
        st.info("💡 **Dùng Power Query:** Thêm `/?export=csv` vào cuối đường dẫn trang web này để dán vào Excel (*Data -> From Web*).")

# ---------------------------------------------------------
# TAB 2: NHẬP KHO (INBOUND) - Sửa hiển thị Tên hàng/Mã hàng
# ---------------------------------------------------------
with tab2:
    st.subheader("📥 THÔNG TIN PHIẾU NHẬP KHO MRO")
    
    with st.form("form_nhap_kho", clear_on_submit=False):
        col_in1, col_in2 = st.columns(2)
        
        with col_in1:
            if not df_inventory.empty:
                # Danh sách dạng "Mã Item - Tên Item"
                item_list = df_inventory.apply(lambda r: f"{r['item_num']} - {r.get('item_name', '')}", axis=1).tolist()
                selected_item_str = st.selectbox("Mã Item# (*)", options=item_list)
                
                # Trích xuất Mã Item
                selected_item_num = selected_item_str.split(" - ")[0]
                
                # Tra cứu thông tin Tên hàng & Tồn kho
                matched_row = df_inventory[df_inventory['item_num'] == selected_item_num].iloc[0]
                item_name_val = matched_row.get('item_name', 'N/A')
                current_stock = matched_row.get('ton_kho', 0)
                
                # HIỂN THỊ TÊN HÀNG & TỒN KHO TRỰC QUAN
                st.info(f"📌 **Tên hàng:** {item_name_val}\n\n📊 **Tồn kho hiện tại:** {current_stock} Pcs")
            else:
                selected_item_num = st.text_input("Mã Item# (*)", value="")
                st.warning("Chưa có danh mục hàng trong CSDL!")
                
            so_luong_nhap = st.number_input("Số lượng nhập (*)", min_value=1, value=1, step=1)

        with col_in2:
            nguoi_nhap = st.text_input("Người thực hiện / Nhân viên Kho")
            location_note = st.text_input("Vị trí lưu kho (Location) / Ghi chú")

        st.markdown("<br>", unsafe_allow_html=True)
        btn_nhap = st.form_submit_button("💾 XÁC NHẬN NHẬP KHO", use_container_width=True)

        if btn_nhap:
            if selected_item_num:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("UPDATE inventory SET ton_kho = ton_kho + ? WHERE item_num = ?", (so_luong_nhap, selected_item_num))
                conn.commit()
                conn.close()
                st.success(f"✅ Đã nhập thành công {so_luong_nhap} Pcs cho mã hàng: {selected_item_num}")
                st.rerun()

# ---------------------------------------------------------
# TAB 3: XUẤT KHO (OUTBOUND)
# ---------------------------------------------------------
with tab3:
    st.subheader("📤 THÔNG TIN PHIẾU XUẤT KHO MRO")
    
    with st.form("form_xuat_kho", clear_on_submit=False):
        col_out1, col_out2 = st.columns(2)
        
        with col_out1:
            if not df_inventory.empty:
                item_list_out = df_inventory.apply(lambda r: f"{r['item_num']} - {r.get('item_name', '')}", axis=1).tolist()
                selected_item_out_str = st.selectbox("Mã Item# (*)", options=item_list_out, key="sb_out")
                selected_item_num_out = selected_item_out_str.split(" - ")[0]
                
                matched_row_out = df_inventory[df_inventory['item_num'] == selected_item_num_out].iloc[0]
                item_name_out_val = matched_row_out.get('item_name', 'N/A')
                current_stock_out = matched_row_out.get('ton_kho', 0)
                
                st.info(f"📌 **Tên hàng:** {item_name_out_val}\n\n📊 **Tồn kho hiện tại:** {current_stock_out} Pcs")
            else:
                selected_item_num_out = st.text_input("Mã Item# (*)", value="", key="ti_out")
                st.warning("Chưa có danh mục hàng trong CSDL!")
                
            so_luong_xuat = st.number_input("Số lượng xuất (*)", min_value=1, value=1, step=1)

        with col_out2:
            nguoi_xuat = st.text_input("Người nhận / Bộ phận yêu cầu")
            ghi_chu_xuat = st.text_input("Ghi chú xuất kho")

        st.markdown("<br>", unsafe_allow_html=True)
        btn_xuat = st.form_submit_button("📤 XÁC NHẬN XUẤT KHO", use_container_width=True)

        if btn_xuat:
            if selected_item_num_out:
                if current_stock_out >= so_luong_xuat:
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("UPDATE inventory SET ton_kho = ton_kho - ? WHERE item_num = ?", (so_luong_xuat, selected_item_num_out))
                    conn.commit()
                    conn.close()
                    st.success(f"✅ Đã xuất thành công {so_luong_xuat} Pcs cho mã hàng: {selected_item_num_out}")
                    st.rerun()
                else:
                    st.error("❌ Số lượng tồn kho không đủ để xuất!")

# ---------------------------------------------------------
# TAB 4: LỊCH SỬ GIAO DỊCH
# ---------------------------------------------------------
with tab4:
    st.subheader("📜 Lịch Sử Giao Dịch Nhập / Xuất Kho")
    st.info("Tính năng xem lại toàn bộ nhật ký giao dịch chi tiết.")