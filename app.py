import io
import sqlite3
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG (Thu gọn vừa đủ, không full màn hình)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Quản Lý Kho MRO",
    page_icon="⚙️",
    layout="centered",  # Giúp giao diện nằm gọn ở giữa màn hình
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# 2. KHỞI TẠO VÀ KẾT NỐI CƠ SỞ DỮ LIỆU (SQLITE)
# ---------------------------------------------------------
DB_FILE = "inventory.db"

def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    # Tạo bảng nếu chưa có
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_num TEXT NOT NULL,
            ton_kho INTEGER DEFAULT 0,
            min_safety INTEGER DEFAULT 0
        )
    """)
    
    # Kiểm tra và tự động thêm cột min_safety nếu CSDL cũ chưa có
    cursor.execute("PRAGMA table_info(inventory)")
    columns = [col[1] for col in cursor.fetchall()]
    if "min_safety" not in columns:
        cursor.execute("ALTER TABLE inventory ADD COLUMN min_safety INTEGER DEFAULT 0")
    
    conn.commit()
    conn.close()

# Chạy khởi tạo database
init_db()

# ---------------------------------------------------------
# 3. LẤY DỮ LIỆU TỪ CSDL
# ---------------------------------------------------------
def load_data():
    conn = get_connection()
    df = pd.read_sql_query("SELECT item_num, ton_kho, min_safety FROM inventory", conn)
    conn.close()
    return df

df_inv_summary = load_data()

# ---------------------------------------------------------
# 4. TÍNH NĂNG POWER QUERY KHÔNG CẦN GIAO DIỆN
# (Khi gọi link: https://your-app.streamlit.app/?export=csv)
# ---------------------------------------------------------
query_params = st.query_params

if query_params.get("export") == "csv":
    # Trả về thuần văn bản CSV để Excel Power Query đọc trực tiếp
    csv_string = df_inv_summary.to_csv(index=False)
    st.text(csv_string)
    st.stop()  # Dừng ứng dụng tại đây, không hiển thị UI

# ---------------------------------------------------------
# 5. GIAO DIỆN TỔNG QUAN HỆ THỐNG (UI)
# ---------------------------------------------------------
# Banner Tiêu đề đã loại bỏ chữ "CHUẨN TPS"
st.markdown("""
    <div style="background-color: #0e1117; padding: 18px 25px; border-radius: 10px; margin-bottom: 25px; color: white;">
        <h2 style="margin:0; font-size: 24px; color: #ffffff;">⚙️ HỆ THỐNG QUẢN LÝ KHO MRO</h2>
        <p style="margin:5px 0 0 0; color: #8a99ad; font-size: 13px;">
            QUẢN LÝ TRỰC QUAN | BẢNG KANBAN KHO PRODUCTION
        </p>
    </div>
""", unsafe_allow_html=True)

# Hiển thị dữ liệu bảng tồn kho
st.subheader("📋 Bảng Tổng Hop Tồn Kho")
st.dataframe(df_inv_summary, use_container_width=True)

# ---------------------------------------------------------
# 6. KHU VỰC TẢI DỮ LIỆU & LINK POWER QUERY
# ---------------------------------------------------------
st.markdown("---")
st.subheader("🔗 Kết nối & Tải dữ liệu")

col1, col2 = st.columns(2)

# Nút tải file Excel trực tiếp
with col1:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_inv_summary.to_excel(writer, index=False, sheet_name='Inventory')
    excel_data = output.getvalue()

    st.download_button(
        label="📥 Tải Báo Cáo Excel (.xlsx)",
        data=excel_data,
        file_name="mro_inventory_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Hướng dẫn & Link Power Query cho Excel
with col2:
    st.info("💡 **Dùng Power Query:** Thêm `/?export=csv` vào cuối đường dẫn trang web này để dán vào Excel (*Data -> From Web*).")