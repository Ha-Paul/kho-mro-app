import io
import sqlite3
from datetime import datetime
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
# 2. KHỞI TẠO VÀ KẾT NỐI FILE CSDL (mro_production.db)
# =========================================================
DB_FILE = "mro_production.db"

def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def check_and_update_schema():
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_num TEXT UNIQUE NOT NULL,
                item_name TEXT DEFAULT '',
                name_vie TEXT DEFAULT '',
                ton_kho INTEGER DEFAULT 0,
                min_safety INTEGER DEFAULT 0,
                location TEXT DEFAULT ''
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                type TEXT,
                item_num TEXT,
                item_name TEXT,
                name_vie TEXT,
                quantity INTEGER,
                operator TEXT,
                note TEXT
            )
        """)
        conn.commit()
    except Exception as e:
        st.error(f"Lỗi cấu trúc CSDL: {e}")
    finally:
        conn.close()

check_and_update_schema()

# =========================================================
# 3. HÀM TỰ ĐỘNG BẮT TÊN CỘT THÔNG MINH
# =========================================================
def get_item_names(row):
    """Trích xuất Tên ENG và Tên VIE từ bất kỳ tên cột nào có trong DB"""
    eng = ""
    vie = ""
    
    # Tìm cột Tên Tiếng Anh
    for col in ['item_name', 'name_eng', 'description', 'eng_name', 'ten_eng']:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != "":
            eng = str(row[col])
            break
            
    # Tìm cột Tên Tiếng Việt
    for col in ['name_vie', 'ten_vie', 'vietnamese_name', 'description_vie', 'ten_hang']:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip() != "":
            vie = str(row[col])
            break
            
    # Nếu không thấy VIE thì lấy tạm ENG làm tên chung và ngược lại
    if not eng and vie: eng = vie
    if not vie and eng: vie = eng
    
    return eng, vie

# =========================================================
# 4. KẾT NỐI POWER QUERY DÙNG CHO EXCEL (Xử lý ngầm)
# =========================================================
query_params = st.query_params

if query_params.get("export") == "csv":
    conn = get_connection()
    df_export = pd.read_sql_query("SELECT * FROM inventory", conn)
    conn.close()
    st.text(df_export.to_csv(index=False))
    st.stop()

# =========================================================
# 5. TRUY VẤN DỮ LIỆU TỔNG QUAN
# =========================================================
conn = get_connection()
try:
    df_inventory = pd.read_sql_query("SELECT * FROM inventory", conn)
except Exception:
    df_inventory = pd.DataFrame()

try:
    df_history = pd.read_sql_query("SELECT * FROM history ORDER BY id DESC", conn)
except Exception:
    df_history = pd.DataFrame()
conn.close()

total_sku = len(df_inventory)
low_stock = 0
if not df_inventory.empty and 'ton_kho' in df_inventory.columns:
    min_col = 'min_safety' if 'min_safety' in df_inventory.columns else 'ton_kho'
    low_stock = len(df_inventory[df_inventory['ton_kho'] <= df_inventory[min_col]])

# Tính tổng nhập/xuất trong ngày
today_str = datetime.now().strftime("%Y-%m-%d")
today_in, today_out = 0, 0
if not df_history.empty and 'timestamp' in df_history.columns:
    df_today = df_history[df_history['timestamp'].str.startswith(today_str, na=False)]
    today_in = df_today[df_today['type'] == 'NHẬP']['quantity'].sum() if not df_today.empty else 0
    today_out = df_today[df_today['type'] == 'XUẤT']['quantity'].sum() if not df_today.empty else 0

# =========================================================
# 6. GIAO DIỆN CHÍNH (BANNER & KPI)
# =========================================================
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
    st.markdown(f"""
        <div style="border-left: 4px solid #1a7f37; background: #ffffff; padding: 12px 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <div style="font-size: 11px; color: #57606a; font-weight: bold; text-transform: uppercase;">NHẬP TRONG NGÀY</div>
            <div style="font-size: 22px; font-weight: bold; color: #1a7f37; margin-top: 4px;">+{today_in} <span style="font-size: 13px; color: #57606a; font-weight: normal;">Pcs</span></div>
        </div>
    """, unsafe_allow_html=True)

with col_kpi4:
    st.markdown(f"""
        <div style="border-left: 4px solid #d97706; background: #ffffff; padding: 12px 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <div style="font-size: 11px; color: #57606a; font-weight: bold; text-transform: uppercase;">XUẤT TRONG NGÀY</div>
            <div style="font-size: 22px; font-weight: bold; color: #d97706; margin-top: 4px;">-{today_out} <span style="font-size: 13px; color: #57606a; font-weight: normal;">Pcs</span></div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =========================================================
# 7. QUẢN LÝ TAB BẢNG ĐIỀU KHIỂN
# =========================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 TỒN KHO REALTIME & KANBAN", 
    "📥 NHẬP KHO (INBOUND)", 
    "📤 XUẤT KHO (OUTBOUND)", 
    "📜 LỊCH SỬ GIAO DỊCH"
])

# ---------------------------------------------------------
# TAB 1: TỒN KHO REALTIME + TÌM KIẾM
# ---------------------------------------------------------
with tab1:
    st.subheader("📋 Bảng Tổng Hợp Tồn Kho")
    search_stock = st.text_input("🔍 Tìm kiếm Mã hàng, Tên ENG hoặc Tên VIE...", key="search_stock")
    
    df_stock_display = df_inventory.copy()
    if search_stock and not df_stock_display.empty:
        # Tìm kiếm không phân biệt chữ hoa/thường trên mọi cột
        mask = df_stock_display.astype(str).apply(lambda x: x.str.contains(search_stock, case=False, na=False)).any(axis=1)
        df_stock_display = df_stock_display[mask]
    
    st.dataframe(df_stock_display, use_container_width=True)
    
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
            file_name="mro_production_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with col_d2:
        st.info("💡 **Dùng Power Query:** Thêm `/?export=csv` vào cuối đường dẫn trang web này để dán vào Excel (*Data -> From Web*).")

# ---------------------------------------------------------
# TAB 2: NHẬP KHO (INBOUND)
# ---------------------------------------------------------
with tab2:
    st.subheader("📥 THÔNG TIN PHIẾU NHẬP KHO MRO")
    
    with st.form("form_nhap_kho", clear_on_submit=False):
        col_in1, col_in2 = st.columns(2)
        
        with col_in1:
            selected_item_num = st.text_input("Mã Item# (*)", value="Sbd0001", help="Nhập mã item để tra cứu").strip()
            
            name_eng_val = ""
            name_vie_val = ""
            current_stock = 0
            is_existing = False
            
            if selected_item_num and not df_inventory.empty:
                # TÌM KIẾM KHÔNG PHÂN BIỆT HOA THƯỜNG (Case-insensitive)
                match = df_inventory[df_inventory['item_num'].astype(str).str.lower() == selected_item_num.lower()]
                
                if not match.empty:
                    matched_row = match.iloc[0]
                    actual_item_num = matched_row['item_num']
                    name_eng_val, name_vie_val = get_item_names(matched_row)
                    current_stock = matched_row.get('ton_kho', 0)
                    is_existing = True
                    
                    st.info(f"📌 **Mã Hàng:** {actual_item_num}\n\n🇬🇧 **Tên ENG:** {name_eng_val if name_eng_val else 'Chưa có thông tin'}\n\n🇻🇳 **Tên VIE:** {name_vie_val if name_vie_val else 'Chưa có thông tin'}\n\n📊 **Tồn kho hiện tại:** {current_stock} Pcs")
            
            if selected_item_num and not is_existing:
                st.warning("✨ Mã Item này chưa có trong CSDL! Vui lòng nhập thông tin mã mới:")
                name_eng_val = st.text_input("Tên Tiếng Anh (Name ENG) (*)", value="")
                name_vie_val = st.text_input("Tên Tiếng Việt (Name VIE) (*)", value="")
                
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
                
                if is_existing:
                    cursor.execute("UPDATE inventory SET ton_kho = ton_kho + ? WHERE LOWER(item_num) = LOWER(?)", (so_luong_nhap, selected_item_num))
                else:
                    cursor.execute("""
                        INSERT INTO inventory (item_num, item_name, name_vie, ton_kho, min_safety, location)
                        VALUES (?, ?, ?, ?, 5, ?)
                    """, (selected_item_num, name_eng_val, name_vie_val, so_luong_nhap, location_note))
                
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                    INSERT INTO history (timestamp, type, item_num, item_name, name_vie, quantity, operator, note)
                    VALUES (?, 'NHẬP', ?, ?, ?, ?, ?, ?)
                """, (now_str, selected_item_num, name_eng_val, name_vie_val, so_luong_nhap, nguoi_nhap, location_note))
                
                conn.commit()
                conn.close()
                st.success(f"✅ Đã nhập thành công {so_luong_nhap} Pcs cho mã hàng: **{selected_item_num}**!")
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
                # Tạo danh sách chọn linh hoạt
                item_options = []
                for _, row in df_inventory.iterrows():
                    eng, vie = get_item_names(row)
                    item_options.append(f"{row['item_num']} | ENG: {eng} | VIE: {vie}")
                    
                selected_item_out_str = st.selectbox("Mã Item# (*)", options=item_options, key="sb_out")
                selected_item_num_out = selected_item_out_str.split(" | ")[0]
                
                matched_row_out = df_inventory[df_inventory['item_num'] == selected_item_num_out].iloc[0]
                name_eng_out, name_vie_out = get_item_names(matched_row_out)
                current_stock_out = matched_row_out.get('ton_kho', 0)
                
                st.info(f"📌 **Mã Hàng:** {selected_item_num_out}\n\n🇬🇧 **Tên ENG:** {name_eng_out}\n\n🇻🇳 **Tên VIE:** {name_vie_out}\n\n📊 **Tồn kho hiện tại:** {current_stock_out} Pcs")
            else:
                selected_item_num_out = st.text_input("Mã Item# (*)", value="", key="ti_out")
                name_eng_out, name_vie_out = "", ""
                current_stock_out = 0
                st.warning("Chưa có dữ liệu danh mục kho!")
                
            so_luong_xuat = st.number_input("Số lượng xuất (*)", min_value=1, value=1, step=1)

        with col_out2:
            nguoi_xuat = st.text_input("Người nhận / Bộ phận yêu cầu")
            ghi_chu_xuat = st.text_input("Ghi chú xuất kho")

        st.markdown("<br>", unsafe_allow_html=True)
        btn_xuat = st.form_submit_button("📤 XÁC NHẬN XUẤT KHO", use_container_width=True)

        if btn_xuat:
            if selected_item_num_out and not df_inventory.empty:
                if current_stock_out >= so_luong_xuat:
                    conn = get_connection()
                    cursor = conn.cursor()
                    
                    cursor.execute("UPDATE inventory SET ton_kho = ton_kho - ? WHERE item_num = ?", (so_luong_xuat, selected_item_num_out))
                    
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    cursor.execute("""
                        INSERT INTO history (timestamp, type, item_num, item_name, name_vie, quantity, operator, note)
                        VALUES (?, 'XUẤT', ?, ?, ?, ?, ?, ?)
                    """, (now_str, selected_item_num_out, name_eng_out, name_vie_out, so_luong_xuat, nguoi_xuat, ghi_chu_xuat))
                    
                    conn.commit()
                    conn.close()
                    st.success(f"✅ Đã xuất thành công {so_luong_xuat} Pcs cho mã hàng: {selected_item_num_out}")
                    st.rerun()
                else:
                    st.error("❌ Số lượng tồn kho không đủ để xuất!")

# ---------------------------------------------------------
# TAB 4: LỊCH SỬ GIAO DỊCH + BỘ LỌC TÌM KIẾM
# ---------------------------------------------------------
with tab4:
    st.subheader("📜 Lịch Sử Giao Dịch Nhập / Xuất Kho")
    
    search_history = st.text_input("🔍 Tìm kiếm lịch sử...", key="search_history")
    
    df_his_display = df_history.copy()
    if search_history and not df_his_display.empty:
        mask = df_his_display.astype(str).apply(lambda x: x.str.contains(search_history, case=False, na=False)).any(axis=1)
        df_his_display = df_his_display[mask]
        
    st.dataframe(df_his_display, use_container_width=True)