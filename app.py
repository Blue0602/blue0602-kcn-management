import re
from pathlib import Path

import pandas as pd
import streamlit as st


# =========================
# 1. Cấu hình trang
# =========================
st.set_page_config(
    page_title="Dashboard Khách hàng KCN",
    page_icon="🏭",
    layout="wide",
)


# =========================
# 2. Cấu hình dữ liệu
# =========================
DEFAULT_FILE = "DS KH KCN - LT.xlsx"

# Các sheet rà soát chỉ có 4 cột, không cùng cấu trúc với bảng dữ liệu chính nên bỏ qua.
IGNORE_SHEET_KEYWORDS = ["ra soat", "rà soát"]

# Đổi tên sheet thành tên KCN dễ đọc hơn.
SHEET_TO_KCN = {
    "KCN LA-BS": "KCN Lộc An - Bình Sơn",
    "KCN LT": "KCN Long Thành",
    "KCN GD": "KCN Gò Dầu",
    "KCN AMATA LT": "KCN Amata Long Thành",
    "LT con lai": "Long Thành còn lại",
}

# Nếu muốn chỉ lấy đúng các sheet KCN, giữ danh sách này.
# Nếu muốn lấy cả DS TONG, thêm "DS TONG" vào danh sách bên dưới.
PREFERRED_SHEETS = [
    "KCN LA-BS",
    "KCN LT",
    "KCN GD",
    "KCN AMATA LT",
    "LT con lai",
]

# Cột cần hiển thị trong bảng chi tiết.
DISPLAY_COLUMNS = [
    "mst",
    "ten_kh",
    "diachi_kh",
    "tong_sl",
    "tong_dthu",
    "ten_kcn",
]

# Các cột sản lượng dịch vụ dùng để xác định khách hàng đang dùng từng loại dịch vụ.
# Bạn có thể bổ sung hoặc đổi tên hiển thị tại đây.
SERVICE_COLUMNS = {
    "Internet / Net": "net_sl",
    "Internet khác": "int_sl",
    "Điện thoại cố định": "dtcd_sl",
    "Truyền số liệu": "tsl_sl",
    "VPN": "vpn_sl",
    "MyTV": "mytv_sl",
    "Chữ ký số": "CA_sl",
    "Hóa đơn điện tử": "HDDT_sl",
    "BHXH": "BHXH_sl",
    "Tracking": "tracking_sl",
    "Brandname": "Brandname_sl",
    "VNPTS": "vnpts_sl",
}


# =========================
# 3. Hàm xử lý dữ liệu
# =========================
def normalize_column_name(col) -> str:
    """Chuẩn hóa tên cột để tránh lỗi khoảng trắng/ký tự thừa."""
    col = str(col).strip()
    col = re.sub(r"\s+", "_", col)
    return col


def find_header_row(excel_file, sheet_name: str, max_scan_rows: int = 10):
    """
    Tự tìm dòng tiêu đề bằng cách quét các dòng đầu sheet.
    File này có sheet DS TONG: header ở dòng 3 trong Excel, tức index 2.
    Nhiều sheet KCN khác có header ở dòng đầu tiên, tức index 0.
    """
    preview = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=None,
        nrows=max_scan_rows,
        engine="openpyxl",
    )

    for idx, row in preview.iterrows():
        values = [str(x).strip().lower() for x in row.tolist() if pd.notna(x)]
        if "mst" in values and "ten_kh" in values:
            return idx

    return None


def infer_kcn_name(sheet_name: str) -> str:
    """Suy luận tên KCN từ tên sheet."""
    return SHEET_TO_KCN.get(sheet_name, sheet_name)


def should_ignore_sheet(sheet_name: str) -> bool:
    """Bỏ qua các sheet không phải bảng dữ liệu chuẩn."""
    lowered = sheet_name.lower().strip()
    return any(keyword in lowered for keyword in IGNORE_SHEET_KEYWORDS)


def clean_dataframe(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """Làm sạch dữ liệu sau khi đọc từng sheet."""
    df = df.copy()

    # Chuẩn hóa tên cột
    df.columns = [normalize_column_name(c) for c in df.columns]

    # Loại bỏ các cột Unnamed nếu có
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False, na=False)]

    # Chỉ xử lý sheet có cột tối thiểu
    required_cols = {"mst", "ten_kh", "tong_sl"}
    if not required_cols.issubset(set(df.columns)):
        return pd.DataFrame()

    # Thêm cột KCN dựa trên tên sheet
    df["source_sheet"] = sheet_name
    df["ten_kcn"] = infer_kcn_name(sheet_name)

    # Loại bỏ dòng tổng cộng hoặc dòng trống
    df["mst"] = df["mst"].astype(str).str.strip()
    df = df[~df["mst"].str.lower().isin(["nan", "none", "", "tổng cộng:", "tong cong:", "tổng cộng"])]
    df = df[~df["mst"].str.contains("tổng", case=False, na=False)]

    # Chuẩn hóa số liệu
    numeric_cols = [c for c in df.columns if c.endswith("_sl") or c.endswith("_dthu")]
    numeric_cols += ["tong_sl", "tong_dthu", "dthu"]
    numeric_cols = list(dict.fromkeys([c for c in numeric_cols if c in df.columns]))

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Chuẩn hóa text
    for col in ["ten_kh", "diachi_kh", "huyen", "phuong", "tinh"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    # Trạng thái sử dụng dịch vụ
    df["trang_thai"] = df["tong_sl"].apply(
        lambda x: "Đang dùng" if pd.notna(x) and float(x) > 0 else "Chưa dùng / Tiềm năng"
    )

    return df


@st.cache_data(show_spinner="Đang tải và xử lý dữ liệu...")
def load_excel_data(file_source) -> pd.DataFrame:
    """
    Đọc file Excel, tự tìm header từng sheet, gom các sheet KCN thành một DataFrame.
    Có dùng @st.cache_data để lần sau load nhanh hơn.
    """
    excel = pd.ExcelFile(file_source, engine="openpyxl")
    frames = []

    for sheet in excel.sheet_names:
        if should_ignore_sheet(sheet):
            continue

        # Ưu tiên lấy các sheet KCN. DS TONG là bảng tổng, dễ gây trùng dữ liệu nếu gom chung.
        if sheet not in PREFERRED_SHEETS:
            continue

        header_row = find_header_row(file_source, sheet)
        if header_row is None:
            continue

        temp = pd.read_excel(
            file_source,
            sheet_name=sheet,
            header=header_row,
            engine="openpyxl",
        )

        temp = clean_dataframe(temp, sheet)
        if not temp.empty:
            frames.append(temp)

    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)

    # Xóa trùng theo mã số thuế + KCN nếu có dòng lặp.
    if "mst" in df.columns and "ten_kcn" in df.columns:
        df = df.drop_duplicates(subset=["mst", "ten_kcn"], keep="first")

    return df


def initialize_session_data(base_df: pd.DataFrame, data_source_key: str) -> None:
    """
    Khởi tạo dữ liệu làm việc trong session_state.
    Khi đổi file nguồn, app tự reset lại dữ liệu theo file mới.
    """
    if (
        "customer_df" not in st.session_state
        or st.session_state.get("data_source_key") != data_source_key
    ):
        st.session_state.customer_df = base_df.copy()
        st.session_state.data_source_key = data_source_key
        st.session_state.added_customers_count = 0


def build_new_customer_row(
    current_df: pd.DataFrame,
    mst: str,
    ten_kh: str,
    ten_kcn: str,
    selected_services: list[str],
    estimated_revenue: float,
) -> dict:
    """
    Tạo một dòng dữ liệu mới theo đúng cấu trúc cột hiện tại của DataFrame.
    Các cột không nhập sẽ được điền giá trị mặc định để tránh lỗi concat.
    """
    numeric_cols = [
        c
        for c in current_df.columns
        if c.endswith("_sl") or c.endswith("_dthu") or c in ["tong_sl", "tong_dthu", "dthu"]
    ]

    new_row = {}
    for col in current_df.columns:
        if col in numeric_cols:
            new_row[col] = 0
        else:
            new_row[col] = ""

    new_row["mst"] = mst.strip()
    new_row["ten_kh"] = ten_kh.strip()
    new_row["diachi_kh"] = ""
    new_row["ten_kcn"] = ten_kcn
    new_row["source_sheet"] = "Nhập liệu thủ công"

    # Mỗi dịch vụ được chọn sẽ được tính là 1 sản lượng để dashboard nhận diện là đang dùng.
    for service_name in selected_services:
        service_sl_col = SERVICE_COLUMNS.get(service_name)
        if service_sl_col and service_sl_col in current_df.columns:
            new_row[service_sl_col] = 1

    total_service_count = len(selected_services)
    new_row["tong_sl"] = total_service_count
    if "tong_dthu" in current_df.columns:
        new_row["tong_dthu"] = float(estimated_revenue)
    if "dthu" in current_df.columns:
        new_row["dthu"] = float(estimated_revenue)

    # Nếu có cột doanh thu theo dịch vụ, chia đều doanh thu ước tính cho các dịch vụ được chọn.
    if selected_services and estimated_revenue > 0:
        revenue_per_service = float(estimated_revenue) / len(selected_services)
        for service_name in selected_services:
            service_sl_col = SERVICE_COLUMNS.get(service_name)
            if not service_sl_col:
                continue
            service_revenue_col = service_sl_col.replace("_sl", "_dthu")
            if service_revenue_col in current_df.columns:
                new_row[service_revenue_col] = revenue_per_service

    new_row["trang_thai"] = "Đang dùng" if total_service_count > 0 else "Chưa dùng / Tiềm năng"

    return new_row


def get_service_usage(df: pd.DataFrame) -> pd.DataFrame:
    """Tính số lượng khách hàng đang dùng từng dịch vụ cụ thể."""
    rows = []
    active_df = df[df["tong_sl"] > 0].copy()

    for service_name, col in SERVICE_COLUMNS.items():
        if col in active_df.columns:
            count = int((active_df[col].fillna(0) > 0).sum())
            rows.append({"Dịch vụ": service_name, "Số khách hàng": count})

    service_df = pd.DataFrame(rows)
    if service_df.empty:
        return service_df

    service_df = service_df.sort_values("Số khách hàng", ascending=False)
    return service_df


def prepare_display_table(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Chuẩn bị bảng hiển thị, chỉ lấy cột tồn tại trong dữ liệu."""
    available_cols = [c for c in columns if c in df.columns]
    result = df[available_cols].copy()

    rename_map = {
        "mst": "Mã số thuế",
        "ten_kh": "Tên khách hàng",
        "diachi_kh": "Địa chỉ",
        "tong_sl": "Tổng sản lượng",
        "tong_dthu": "Tổng doanh thu",
        "ten_kcn": "Khu công nghiệp",
    }
    result = result.rename(columns=rename_map)
    return result


def convert_df_to_csv(df: pd.DataFrame) -> bytes:
    """Chuyển DataFrame sang CSV để tải về."""
    return df.to_csv(index=False).encode("utf-8-sig")


# =========================
# 4. Giao diện chính
# =========================
st.title("🏭 Dashboard quản lý khách hàng doanh nghiệp tại KCN")
st.caption("Lọc KCN, phân nhóm khách hàng đang dùng/chưa dùng dịch vụ và xem thống kê dịch vụ chi tiết.")

with st.sidebar:
    st.header("⚙️ Cấu hình dữ liệu")
    uploaded_file = st.file_uploader(
        "Tải file Excel dữ liệu KCN",
        type=["xlsx"],
        help="Nếu không tải file, app sẽ tự tìm file DS KH KCN - LT.xlsx trong thư mục dự án.",
    )

    st.markdown("---")
    st.info(
        "App đang tự lấy tên KCN từ tên sheet như KCN LT, KCN GD, KCN LA-BS, KCN AMATA LT."
    )

# Ưu tiên file người dùng upload. Nếu không có, đọc file mặc định trong repo.
if uploaded_file is not None:
    file_source = uploaded_file
    data_source_key = f"uploaded::{uploaded_file.name}::{uploaded_file.size}"
else:
    file_path = Path(DEFAULT_FILE)
    if not file_path.exists():
        st.error(
            f"Không tìm thấy file `{DEFAULT_FILE}` trong thư mục dự án. "
            "Hãy upload file ở thanh bên trái hoặc đưa file Excel vào cùng thư mục với app.py."
        )
        st.stop()
    file_source = str(file_path)
    data_source_key = f"default::{DEFAULT_FILE}"

# Đọc dữ liệu gốc và đưa vào session_state để có thể thêm dòng mới ngay trên web.
base_df = load_excel_data(file_source)

if base_df.empty:
    st.error(
        "Không đọc được dữ liệu chuẩn. Hãy kiểm tra file có các cột `mst`, `ten_kh`, `tong_sl` hay không."
    )
    st.stop()

initialize_session_data(base_df, data_source_key)
df = st.session_state.customer_df.copy()

# Danh sách KCN/dịch vụ dùng chung cho form và bộ lọc.
kcn_list = sorted(df["ten_kcn"].dropna().unique().tolist())
service_options = [name for name, col in SERVICE_COLUMNS.items() if col in df.columns]


# =========================
# 5. Form thêm khách hàng
# =========================
with st.sidebar:
    st.markdown("---")
    st.header("➕ Thêm Khách Hàng")

    with st.form("add_customer_form", clear_on_submit=True):
        input_mst = st.text_input("Mã số thuế", placeholder="Ví dụ: 3600123456")
        input_ten_kh = st.text_input("Tên doanh nghiệp", placeholder="Ví dụ: Công ty TNHH ABC")
        input_kcn = st.selectbox("Khu Công Nghiệp", kcn_list)
        input_services = st.multiselect(
            "Dịch vụ sử dụng",
            service_options,
            help="Chọn một hoặc nhiều dịch vụ. Nếu không chọn dịch vụ, khách hàng sẽ được tính là nhóm tiềm năng/chưa dùng.",
        )
        input_revenue = st.number_input(
            "Tổng doanh thu ước tính",
            min_value=0.0,
            value=0.0,
            step=100000.0,
            format="%.0f",
        )

        submitted = st.form_submit_button("Thêm Khách Hàng Mới", use_container_width=True)

        if submitted:
            if not input_mst.strip():
                st.error("Vui lòng nhập Mã số thuế.")
            elif not input_ten_kh.strip():
                st.error("Vui lòng nhập Tên doanh nghiệp.")
            else:
                existing_mask = (
                    (st.session_state.customer_df["mst"].astype(str).str.strip() == input_mst.strip())
                    & (st.session_state.customer_df["ten_kcn"].astype(str) == input_kcn)
                )

                if existing_mask.any():
                    st.warning("Mã số thuế này đã tồn tại trong KCN đã chọn. Dữ liệu chưa được thêm để tránh trùng.")
                else:
                    new_row = build_new_customer_row(
                        current_df=st.session_state.customer_df,
                        mst=input_mst,
                        ten_kh=input_ten_kh,
                        ten_kcn=input_kcn,
                        selected_services=input_services,
                        estimated_revenue=input_revenue,
                    )

                    new_row_df = pd.DataFrame([new_row], columns=st.session_state.customer_df.columns)
                    st.session_state.customer_df = pd.concat(
                        [st.session_state.customer_df, new_row_df],
                        ignore_index=True,
                    )
                    st.session_state.added_customers_count += 1
                    st.success("Đã thêm khách hàng mới. Dashboard đang cập nhật lại dữ liệu.")
                    st.rerun()

    if st.session_state.get("added_customers_count", 0) > 0:
        st.caption(f"Đã thêm trong phiên này: {st.session_state.added_customers_count} khách hàng")

    if st.button("🔄 Khôi phục dữ liệu gốc", use_container_width=True):
        st.session_state.customer_df = base_df.copy()
        st.session_state.added_customers_count = 0
        st.success("Đã khôi phục dữ liệu gốc từ file Excel.")
        st.rerun()

# Sau khi form có thể cập nhật session_state, lấy lại DataFrame mới nhất để dashboard tính lại.
df = st.session_state.customer_df.copy()


# =========================
# 6. Bộ lọc KCN
# =========================
st.subheader("🔎 Bộ lọc")

kcn_options = sorted(df["ten_kcn"].dropna().unique().tolist())
kcn_options = ["Tất cả KCN"] + kcn_options

selected_kcn = st.selectbox("Chọn Khu công nghiệp", kcn_options)

if selected_kcn == "Tất cả KCN":
    filtered_df = df.copy()
else:
    filtered_df = df[df["ten_kcn"] == selected_kcn].copy()

# =========================
# 7. KPI tổng quan
# =========================
total_customers = int(filtered_df["mst"].nunique())
active_customers_df = filtered_df[filtered_df["tong_sl"] > 0].copy()
potential_customers_df = filtered_df[(filtered_df["tong_sl"].isna()) | (filtered_df["tong_sl"] <= 0)].copy()

active_customers = int(active_customers_df["mst"].nunique())
potential_customers = int(potential_customers_df["mst"].nunique())
penetration_rate = (active_customers / total_customers * 100) if total_customers > 0 else 0

total_revenue = float(filtered_df["tong_dthu"].sum()) if "tong_dthu" in filtered_df.columns else 0

st.subheader("📌 Thống kê tổng quan")
col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("Tổng số DN", f"{total_customers:,}")
col2.metric("DN đang dùng", f"{active_customers:,}")
col3.metric("DN chưa dùng / tiềm năng", f"{potential_customers:,}")
col4.metric("Tỷ lệ khai thác", f"{penetration_rate:.1f}%")
col5.metric("Tổng doanh thu", f"{total_revenue:,.0f}")

# =========================
# 8. Phân tích dịch vụ chi tiết
# =========================
st.subheader("📊 Phân tích dịch vụ đang sử dụng")
service_usage = get_service_usage(filtered_df)

if service_usage.empty:
    st.warning("Chưa tìm thấy cột sản lượng dịch vụ phù hợp để vẽ biểu đồ.")
else:
    left, right = st.columns([2, 1])

    with left:
        chart_data = service_usage.set_index("Dịch vụ")
        st.bar_chart(chart_data)

    with right:
        st.dataframe(
            service_usage,
            use_container_width=True,
            hide_index=True,
        )

# =========================
# 9. Bảng chi tiết theo tabs
# =========================
st.subheader("📋 Danh sách khách hàng chi tiết")

tab1, tab2, tab3 = st.tabs(["✅ Đang dùng dịch vụ", "🎯 Tiềm năng / Chưa dùng", "➕ Thêm Khách Hàng"])

with tab1:
    st.write(f"Số khách hàng đang dùng dịch vụ: **{active_customers:,}**")
    active_table = prepare_display_table(active_customers_df, DISPLAY_COLUMNS)

    st.dataframe(
        active_table,
        use_container_width=True,
        hide_index=True,
        height=500,
    )

    st.download_button(
        label="⬇️ Tải danh sách đang dùng CSV",
        data=convert_df_to_csv(active_table),
        file_name="khach_hang_dang_dung.csv",
        mime="text/csv",
    )

with tab2:
    st.write(f"Số khách hàng tiềm năng/chưa dùng: **{potential_customers:,}**")
    potential_cols = ["mst", "ten_kh", "diachi_kh", "ten_kcn"]
    potential_table = prepare_display_table(potential_customers_df, potential_cols)

    st.dataframe(
        potential_table,
        use_container_width=True,
        hide_index=True,
        height=500,
    )

    st.download_button(
        label="⬇️ Tải danh sách tiềm năng CSV",
        data=convert_df_to_csv(potential_table),
        file_name="khach_hang_tiem_nang.csv",
        mime="text/csv",
    )

with tab3:
    st.markdown(
        """
        Khu vực thêm khách hàng mới hiện đang nằm ở **thanh Sidebar bên trái**.
        Sau khi bấm **Thêm Khách Hàng Mới**, dashboard sẽ cập nhật ngay các KPI, biểu đồ và bảng danh sách.

        **Lưu ý:** dữ liệu thêm mới hiện chỉ được lưu trong phiên làm việc của Streamlit. Nếu bạn refresh app, redeploy app hoặc server restart, dữ liệu sẽ quay về file Excel gốc. Để lưu vĩnh viễn, bước tiếp theo nên kết nối Supabase/PostgreSQL.
        """
    )

    added_rows = df[df.get("source_sheet", "") == "Nhập liệu thủ công"] if "source_sheet" in df.columns else pd.DataFrame()
    if not added_rows.empty:
        st.write("Các khách hàng đã thêm trong phiên này:")
        added_table = prepare_display_table(added_rows, DISPLAY_COLUMNS)
        st.dataframe(added_table, use_container_width=True, hide_index=True)
    else:
        st.info("Chưa có khách hàng nào được thêm trong phiên này.")

# =========================
# 10. Ghi chú kỹ thuật
# =========================
with st.expander("📝 Ghi chú xử lý dữ liệu"):
    st.markdown(
        """
        - App tự tìm dòng tiêu đề chứa `mst` và `ten_kh`, nên xử lý được sheet có dòng tiêu đề thừa.
        - Các sheet `Ra soat...` không được dùng vì không có cấu trúc cột đầy đủ như bảng chính.
        - Vì file chưa có cột `ten_kcn` riêng, app tự tạo `ten_kcn` dựa trên tên sheet.
        - Điều kiện **Đang dùng dịch vụ**: `tong_sl > 0`.
        - Điều kiện **Tiềm năng / Chưa dùng**: `tong_sl <= 0` hoặc null.
        - Form thêm khách hàng dùng `st.session_state`, nên dữ liệu thêm mới chỉ tồn tại trong phiên chạy hiện tại.
        - Nếu muốn lưu dữ liệu thêm mới vĩnh viễn, nên kết nối database như Supabase/PostgreSQL thay vì ghi ngược vào file Excel trên Streamlit Cloud.
        - Nếu muốn gom thêm sheet `DS TONG`, hãy thêm `"DS TONG"` vào biến `PREFERRED_SHEETS`.
        """
    )
