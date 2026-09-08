from datetime import datetime
import os
import openpyxl
import pandas as pd
import streamlit as st

# ==========================================
# 1. CẤU HÌNH TRANG WEB & GIAO DIỆN CHUẨN TPS
# ==========================================
st.set_page_config(
    page_title="Hệ Thống Quản Lý Sản Xuất & MRO (TPS Standard)",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    div.stButton > button:first-child {
        background-color: #0056b3;
        color: white;
        border-radius: 4px;
        font-weight: bold;
        border: none;
        padding: 0.5rem 1rem;
        width: 100%;
    }
    div.stButton > button:first-child:hover {
        background-color: #004085;
        color: white;
    }
    .metric-card {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #0056b3;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. XỬ LÝ DỮ LIỆU EXCEL (BẢO VỆ DATA HIỆN CÓ)
# ==========================================
EXCEL_FILE = "danh_muc_mro.xlsx"


def load_data():
  """Đọc dữ liệu từ file Excel.

  Nếu file đã có sẵn dữ liệu thì giữ nguyên 100%, chỉ bổ sung các sheet thiếu.
  """
  # 1. Nếu chưa có file Excel nào, tạo mới file mẫu
  if not os.path.exists(EXCEL_FILE):
    df_dm = pd.DataFrame(
        columns=["Item#", "Mã hàng", "Tên ENG", "Tên VIE", "Đơn vị tính"]
    )
    df_nhap = pd.DataFrame(
        columns=[
            "Ngày tháng",
            "Item#",
            "Mã hàng",
            "Tên ENG",
            "Tên VIE",
            "Đơn vị tính",
            "Số lượng nhập",
            "Người nhập",
            "Ghi chú",
        ]
    )
    df_xuat = pd.DataFrame(
        columns=[
            "Ngày tháng",
            "Item#",
            "Mã hàng",
            "Tên ENG",
            "Tên VIE",
            "Đơn vị tính",
            "Số lượng xuất",
            "Người xuất",
            "Ghi chú/Máy",
        ]
    )

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
      df_dm.to_excel(writer, sheet_name="DanhMuc", index=False)
      df_nhap.to_excel(writer, sheet_name="NhapKho", index=False)
      df_xuat.to_excel(writer, sheet_name="XuatKho", index=False)
    return df_dm, df_nhap, df_xuat

  # 2. Nếu file đã tồn tại -> Đọc danh sách sheet để bổ sung nếu thiếu
  wb = openpyxl.load_workbook(EXCEL_FILE)
  sheets = wb.sheetnames

  # Nếu sheet 'DanhMuc' chưa đúng tên (vd: Sheet1), kiểm tra và đổi tên hoặc dùng sheet đầu tiên
  target_sheet = "DanhMuc"
  if "DanhMuc" not in sheets:
    target_sheet = sheets[0]  # Lấy sheet đầu tiên chứa danh mục của bạn

  # Đọc sheet Danh Mục
  df_dm = pd.read_excel(
      EXCEL_FILE, sheet_name=target_sheet, dtype={"Item#": str}
  )

  # Đảm bảo các cột tối thiểu
  for col in ["Item#", "Mã hàng", "Tên ENG", "Tên VIE", "Đơn vị tính"]:
    if col not in df_dm.columns:
      df_dm[col] = ""

  # Tự động tạo sheet NhapKho nếu thiếu
  if "NhapKho" in sheets:
    df_nhap = pd.read_excel(
        EXCEL_FILE, sheet_name="NhapKho", dtype={"Item#": str}
    )
  else:
    df_nhap = pd.DataFrame(
        columns=[
            "Ngày tháng",
            "Item#",
            "Mã hàng",
            "Tên ENG",
            "Tên VIE",
            "Đơn vị tính",
            "Số lượng nhập",
            "Người nhập",
            "Ghi chú",
        ]
    )
    with pd.ExcelWriter(
        EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
      df_nhap.to_excel(writer, sheet_name="NhapKho", index=False)

  # Tự động tạo sheet XuatKho nếu thiếu
  if "XuatKho" in sheets:
    df_xuat = pd.read_excel(
        EXCEL_FILE, sheet_name="XuatKho", dtype={"Item#": str}
    )
  else:
    df_xuat = pd.DataFrame(
        columns=[
            "Ngày tháng",
            "Item#",
            "Mã hàng",
            "Tên ENG",
            "Tên VIE",
            "Đơn vị tính",
            "Số lượng xuất",
            "Người xuất",
            "Ghi chú/Máy",
        ]
    )
    with pd.ExcelWriter(
        EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
      df_xuat.to_excel(writer, sheet_name="XuatKho", index=False)

  wb.close()
  return df_dm, df_nhap, df_xuat


def save_sheet(df, sheet_name):
  """Lưu một DataFrame vào Sheet cụ thể trong file Excel."""
  with pd.ExcelWriter(
      EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace"
  ) as writer:
    df.to_excel(writer, sheet_name=sheet_name, index=False)


df_dm, df_nhap, df_xuat = load_data()

# ==========================================
# 3. GIAO DIỆN CHÍNH (MAIN TABS)
# ==========================================
#st.title("🏭 HỆ THỐNG QUẢN LÝ SẢN XUẤT & LOGISTICS")
st.title("🏭 HỆ THỐNG QUẢN LÝ SẢN XUẤT")

main_tab1, main_tab2 = st.tabs([
   "📊 QUẢN LÝ SẢN XUẤT",
    "⚙️ HỆ THỐNG QUẢN LÝ MRO - FORMING",
])

# ------------------------------------------
# TAB 1: QUẢN LÝ SẢN XUẤT
# ------------------------------------------
with main_tab1:
  st.header("CHƯƠNG TRÌNH QUẢN LÝ SẢN XUẤT TỔNG THỂ")
  st.info(
      "📌 Tab Quản lý sản xuất tổng quan: Theo dõi chỉ số OEE, Tiến độ sản xuất"
      " Real-time, Kế hoạch & Sản lượng."
  )

  col1, col2, col3, col4 = st.columns(4)
  with col1:
    st.metric(
        label="Hiệu suất chung (OEE)", value="85.4%", delta="1.2% (Tăng)"
    )
  with col2:
    st.metric(
        label="Sản lượng Kế hoạch", value="12,000 Pcs", delta="Đúng tiến độ"
    )
  with col3:
    st.metric(
        label="Sản lượng Thực tế", value="10,250 Pcs", delta="-1,750 Pcs"
    )
  with col4:
    st.metric(label="Tỷ lệ Lỗi (NG Rate)", value="0.42%", delta="-0.05% (Tốt)")

  st.divider()
  st.subheader("Trạng thái Dây chuyền Sản xuất (Visual Management)")
  cols_line = st.columns(3)
  lines = [
      {"name": "Line Forming 01", "status": "Đang chạy", "color": "green"},
      {"name": "Line Forming 02", "status": "Đang Bảo trì", "color": "orange"},
      {"name": "Line Forming 03", "status": "Đang chạy", "color": "green"},
  ]
  for idx, line in enumerate(lines):
    with cols_line[idx]:
      st.markdown(f"""
            <div class="metric-card">
                <h4>{line['name']}</h4>
                <p>Trạng thái: <b><span style="color:{line['color']};">{line['status']}</span></b></p>
            </div>
            """, unsafe_allow_html=True)

# ------------------------------------------
# TAB 2: HỆ THỐNG QUẢN LÝ MRO - FORMING (TPS)
# ------------------------------------------
with main_tab2:
  st.header("⚙️ QUẢN LÝ MRO - XƯỞNG FORMING")
  st.caption(
      "Chuẩn hóa nguyên tắc TPS: Loại bỏ lãng phí (Muda) - Tự động hóa"
      " (Jidoka) - Đúng thời điểm (JIT)"
  )

  mro_tab1, mro_tab2, mro_tab3 = st.tabs(
      ["📥 NHẬP KHO", "📤 XUẤT KHO", "📦 TỒN KHO"]
  )

  item_list = df_dm["Item#"].astype(str).dropna().tolist()

  # ------------------------------------------
  # SUB-TAB 1: NHẬP KHO
  # ------------------------------------------
  with mro_tab1:
    st.subheader("Nhập phụ tùng / Vật tư MRO")

    with st.expander("➕ Thêm mới Mã hàng/Item# vào Danh mục MRO"):
      with st.form("form_add_new_item", clear_on_submit=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        new_item = c1.text_input("Item# mới (*)")
        new_ma = c2.text_input("Mã hàng (*)")
        new_eng = c3.text_input("Tên ENG")
        new_vie = c4.text_input("Tên VIE (*)")
        new_dvt = c5.selectbox("Đơn vị tính (*)", options=["PCS", "SET"])

        btn_add = st.form_submit_button("Lưu Mã Hàng Mới")
        if btn_add:
          if not new_item or not new_ma or not new_vie:
            st.error("❌ Vui lòng điền đầy đủ các thông tin bắt buộc (*)")
          elif new_item.strip() in df_dm["Item#"].astype(str).values:
            st.warning(f"⚠️ Item# [{new_item}] đã tồn tại trong Hệ thống!")
          else:
            new_row = pd.DataFrame([{
                "Item#": str(new_item).strip(),
                "Mã hàng": new_ma,
                "Tên ENG": new_eng,
                "Tên VIE": new_vie,
                "Đơn vị tính": new_dvt,
            }])
            df_dm = pd.concat([df_dm, new_row], ignore_index=True)
            save_sheet(df_dm, "DanhMuc")
            st.success(f"✅ Đã thêm thành công Item# [{new_item}]!")
            st.rerun()

    st.divider()

    col_i1, col_i2 = st.columns([1, 2])
    with col_i1:
      selected_item_nhap = st.selectbox(
          "Chọn Item# (*)", options=[""] + item_list, key="select_nhap_item"
      )

    info_ma, info_eng, info_vie, info_dvt = "", "", "", "PCS"
    if selected_item_nhap:
      matched = df_dm[df_dm["Item#"].astype(str) == selected_item_nhap]
      if not matched.empty:
        info_ma = matched.iloc[0].get("Mã hàng", "")
        info_eng = matched.iloc[0].get("Tên ENG", "")
        info_vie = matched.iloc[0].get("Tên VIE", "")
        info_dvt_val = str(matched.iloc[0].get("Đơn vị tính", "PCS")).upper()
        if info_dvt_val in ["PCS", "SET"]:
          info_dvt = info_dvt_val

    st.info(
        f"🔍 **Thông tin Vật tư:** Mã hàng: **{info_ma}** | Tên VIE:"
        f" **{info_vie}** | Tên ENG: **{info_eng}**"
    )

    with st.form("form_nhap_kho", clear_on_submit=False):
      c_n1, c_n2, c_n3, c_n4 = st.columns(4)
      nguoi_nhap = c_n1.text_input("Người nhập (*)")
      so_luong_nhap = c_n2.number_input(
          "Số lượng nhập (*)", min_value=1, step=1
      )
      dvt_nhap = c_n3.selectbox(
          "Đơn vị tính (*)",
          options=["PCS", "SET"],
          index=0 if info_dvt == "PCS" else 1,
      )
      ngay_nhap = c_n4.date_input("Ngày nhập", value=datetime.now())

      ghi_chu_nhap = st.text_area("Ghi chú nhập kho")

      submit_nhap = st.form_submit_button("XÁC NHẬN NHẬP KHO")

      if submit_nhap:
        if not selected_item_nhap:
          st.error("❌ BẮT BUỘC: Vui lòng chọn Item#!")
        elif not nguoi_nhap.strip():
          st.error("❌ BẮT BUỘC: Ô 'Người nhập' không được để trống!")
        else:
          new_entry = pd.DataFrame([{
              "Ngày tháng": ngay_nhap.strftime("%Y-%m-%d"),
              "Item#": str(selected_item_nhap),
              "Mã hàng": info_ma,
              "Tên ENG": info_eng,
              "Tên VIE": info_vie,
              "Đơn vị tính": dvt_nhap,
              "Số lượng nhập": so_luong_nhap,
              "Người nhập": nguoi_nhap.strip(),
              "Ghi chú": ghi_chu_nhap,
          }])
          df_nhap = pd.concat([df_nhap, new_entry], ignore_index=True)
          save_sheet(df_nhap, "NhapKho")
          st.success("✅ Ghi nhận Nhập kho thành công!")
          st.rerun()

  # ------------------------------------------
  # SUB-TAB 2: XUẤT KHO
  # ------------------------------------------
  with mro_tab2:
    st.subheader("Xuất phụ tùng / Vật tư MRO")

    col_x1, col_x2 = st.columns([1, 2])
    with col_x1:
      selected_item_xuat = st.selectbox(
          "Chọn Item# (*)", options=[""] + item_list, key="select_xuat_item"
      )

    info_ma_x, info_eng_x, info_vie_x, info_dvt_x = "", "", "", "PCS"
    if selected_item_xuat:
      matched_x = df_dm[df_dm["Item#"].astype(str) == selected_item_xuat]
      if not matched_x.empty:
        info_ma_x = matched_x.iloc[0].get("Mã hàng", "")
        info_eng_x = matched_x.iloc[0].get("Tên ENG", "")
        info_vie_x = matched_x.iloc[0].get("Tên VIE", "")
        info_dvt_val_x = str(
            matched_x.iloc[0].get("Đơn vị tính", "PCS")
        ).upper()
        if info_dvt_val_x in ["PCS", "SET"]:
          info_dvt_x = info_dvt_val_x

    st.info(
        f"🔍 **Thông tin Vật tư:** Mã hàng: **{info_ma_x}** | Tên VIE:"
        f" **{info_vie_x}** | Tên ENG: **{info_eng_x}**"
    )

    with st.form("form_xuat_kho", clear_on_submit=False):
      c_x1, c_x2, c_x3, c_x4 = st.columns(4)
      nguoi_xuat = c_x1.text_input("Người Xuất (*)")
      so_luong_xuat = c_x2.number_input(
          "Số lượng Xuất (*)", min_value=1, step=1
      )
      dvt_xuat = c_x3.selectbox(
          "Đơn vị tính (*)",
          options=["PCS", "SET"],
          index=0 if info_dvt_x == "PCS" else 1,
      )
      ngay_xuat = c_x4.date_input("Ngày Xuất", value=datetime.now())

      ghi_chu_xuat = st.text_input("Ghi chú / Mã Máy sử dụng")

      submit_xuat = st.form_submit_button("XÁC NHẬN XUẤT KHO")

      if submit_xuat:
        if not selected_item_xuat:
          st.error("❌ BẮT BUỘC: Vui lòng chọn Item#!")
        elif not nguoi_xuat.strip():
          st.error("❌ BẮT BUỘC: Ô 'Người Xuất' không được để trống!")
        else:
          new_exit = pd.DataFrame([{
              "Ngày tháng": ngay_xuat.strftime("%Y-%m-%d"),
              "Item#": str(selected_item_xuat),
              "Mã hàng": info_ma_x,
              "Tên ENG": info_eng_x,
              "Tên VIE": info_vie_x,
              "Đơn vị tính": dvt_xuat,
              "Số lượng xuất": so_luong_xuat,
              "Người xuất": nguoi_xuat.strip(),
              "Ghi chú/Máy": ghi_chu_xuat,
          }])
          df_xuat = pd.concat([df_xuat, new_exit], ignore_index=True)
          save_sheet(df_xuat, "XuatKho")
          st.success("✅ Ghi nhận Xuất kho thành công!")
          st.rerun()

  # ------------------------------------------
  # SUB-TAB 3: TỒN KHO & TÌM KIẾM
  # ------------------------------------------
  with mro_tab3:
    st.subheader("Báo cáo Tồn kho MRO Real-time")

    df_ton = df_dm.copy()

    if not df_nhap.empty and "Số lượng nhập" in df_nhap.columns:
      sum_nhap = (
          df_nhap.groupby("Item#")["Số lượng nhập"].sum().reset_index()
      )
      df_ton = pd.merge(df_ton, sum_nhap, on="Item#", how="left")
      df_ton["Số lượng nhập"] = df_ton["Số lượng nhập"].fillna(0)
    else:
      df_ton["Số lượng nhập"] = 0

    if not df_xuat.empty and "Số lượng xuất" in df_xuat.columns:
      sum_xuat = (
          df_xuat.groupby("Item#")["Số lượng xuất"].sum().reset_index()
      )
      df_ton = pd.merge(df_ton, sum_xuat, on="Item#", how="left")
      df_ton["Số lượng xuất"] = df_ton["Số lượng xuất"].fillna(0)
    else:
      df_ton["Số lượng xuất"] = 0

    df_ton["Tồn kho khả dụng"] = (
        df_ton["Số lượng nhập"] - df_ton["Số lượng xuất"]
    )

    search_kw = st.text_input(
        "🔎 Tìm kiếm linh hoạt (Nhập Item#, Mã hàng, Tên ENG hoặc Tên VIE):", ""
    )

    if search_kw:
      mask = (
          df_ton["Item#"].astype(str).str.contains(search_kw, case=False)
          | df_ton["Mã hàng"].astype(str).str.contains(search_kw, case=False)
          | df_ton["Tên ENG"].astype(str).str.contains(search_kw, case=False)
          | df_ton["Tên VIE"].astype(str).str.contains(search_kw, case=False)
      )
      df_display = df_ton[mask]
    else:
      df_display = df_ton

    cols_to_show = [
        col
        for col in [
            "Item#",
            "Mã hàng",
            "Tên ENG",
            "Tên VIE",
            "Đơn vị tính",
            "Số lượng nhập",
            "Số lượng xuất",
            "Tồn kho khả dụng",
        ]
        if col in df_display.columns
    ]

    st.dataframe(
        df_display[cols_to_show], use_container_width=True, hide_index=True
    )

    try:
      with open(EXCEL_FILE, "rb") as f:
        file_bytes = f.read()

      st.download_button(
          label="📥 Tải Báo Cáo Tồn Kho Excel",
          data=file_bytes,
          file_name=f"Bao_Cao_Ton_Kho_MRO_{datetime.now().strftime('%Y%m%d')}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )
    except Exception as e:
      st.error(f"Không thể tải file báo cáo: {e}")

      # =========================================================
# 3. KẾT NỐI POWER QUERY DÙNG CHO EXCEL (TẠO FILE CSV TĨNH)
# =========================================================
# Mỗi khi app chạy hoặc có thao tác Nhập/Xuất, ghi đè file mro_export.csv
try:
    conn = get_connection()
    df_export = pd.read_sql_query("SELECT * FROM inventory", conn)
    conn.close()
    
    # Ghi file CSV ra thư mục làm việc để Streamlit public hoặc đọc trực tiếp
    df_export.to_csv("mro_export.csv", index=False, encoding='utf-8-sig')
except Exception as e:
    pass

# =========================================================
# 4. KẾT NỐI POWER QUERY DÙNG CHO EXCEL
# =========================================================
query_params = st.query_params

if query_params.get("export") == "csv":
    conn = get_connection()
    try:
        df_export = pd.read_sql_query("SELECT * FROM inventory", conn)
    except Exception:
        df_export = pd.DataFrame()
    finally:
        conn.close()
    
    # Chuyển đổi dữ liệu sang CSV mã hóa UTF-8 chuẩn cho Excel
    csv_data = df_export.to_csv(index=False, encoding='utf-8-sig')
    
    # Hiển thị dữ liệu dạng Code Block thuần để Power Query bóc tách
    st.code(csv_data, language="text")
    st.stop()