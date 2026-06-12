import os
import glob
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# -----------------------------------------------------------------------------
# CONFIG TRANG DASHBOARD
# -----------------------------------------------------------------------------
st.set_page_config(page_title="DMS Sales Dashboard - Bidiphar", layout="wide")

# =============================================================================
# HÀM BUILD PDF (khai báo sớm, dùng được ở mọi nơi)
# =============================================================================
def _build_pdf(file_name, total_val, delivered_val, remain_val, rate_v,
               t_ord, d_ord, r_ord, rate_o,
               df_res, df_asm, df_p, df_detail):
    import os, datetime, unicodedata
    from io import BytesIO
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # Hàm escape XML cho Paragraph (tránh lỗi parse)
    def xesc(s):
        return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

    fn, fb = "Helvetica", "Helvetica-Bold"

    # Download NotoSans hỗ trợ tiếng Việt đầy đủ nếu chưa có
    import tempfile, urllib.request
    _cache = os.path.join(tempfile.gettempdir(), "NotoSans_viet")
    _reg_path  = os.path.join(_cache, "NotoSans-Regular.ttf")
    _bold_path = os.path.join(_cache, "NotoSans-Bold.ttf")
    os.makedirs(_cache, exist_ok=True)

    _noto_reg  = "https://github.com/notofonts/notofonts.github.io/raw/main/fonts/NotoSans/unhinted/ttf/NotoSans-Regular.ttf"
    _noto_bold = "https://github.com/notofonts/notofonts.github.io/raw/main/fonts/NotoSans/unhinted/ttf/NotoSans-Bold.ttf"

    def _dl(url, dest):
        if not os.path.exists(dest):
            try:
                urllib.request.urlretrieve(url, dest)
            except Exception:
                return False
        return os.path.exists(dest)

    if _dl(_noto_reg, _reg_path):
        try:
            pdfmetrics.registerFont(TTFont("NotoSans",     _reg_path))
            fn = "NotoSans"
        except Exception:
            pass
    if _dl(_noto_bold, _bold_path):
        try:
            pdfmetrics.registerFont(TTFont("NotoSans-Bold", _bold_path))
            fb = "NotoSans-Bold"
        except Exception:
            fb = fn

    # Fallback: thử font hệ thống nếu download thất bại
    if fn == "Helvetica":
        _sys_pairs = [
            (os.path.normpath("C:/Windows/Fonts/arial.ttf"),
             os.path.normpath("C:/Windows/Fonts/arialbd.ttf")),
            ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
             "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
            ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
             "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
        ]
        for _r, _b in _sys_pairs:
            if os.path.exists(_r):
                try:
                    pdfmetrics.registerFont(TTFont("_VF", _r))
                    fn = "_VF"
                    if os.path.exists(_b):
                        pdfmetrics.registerFont(TTFont("_VFB", _b))
                        fb = "_VFB"
                    else:
                        fb = "_VF"
                    break
                except Exception:
                    continue

    def ps(sz, bold=False, color="#000000", sb=0, sa=2):
        return ParagraphStyle("_", fontSize=sz,
                               fontName=fb if bold else fn,
                               textColor=colors.HexColor(color),
                               spaceBefore=sb, spaceAfter=sa)

    def mktbl(data, widths, hc):
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0),  colors.HexColor(hc)),
            ("TEXTCOLOR",     (0,0),(-1,0),  colors.white),
            ("FONTNAME",      (0,0),(-1,0),  fb),
            ("FONTNAME",      (0,1),(-1,-1), fn),
            ("FONTSIZE",      (0,0),(-1,-1), 8),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),
             [colors.HexColor("#fdf5f5"), colors.white]),
            ("GRID",          (0,0),(-1,-1), 0.4, colors.HexColor("#cccccc")),
            ("LEFTPADDING",   (0,0),(-1,-1), 5),
            ("RIGHTPADDING",  (0,0),(-1,-1), 5),
            ("TOPPADDING",    (0,0),(-1,-1), 3),
            ("BOTTOMPADDING", (0,0),(-1,-1), 3),
        ]))
        return t

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    story = [
        Paragraph("BÁO CÁO HIỆU SUẤT GIAO HÀNG DMS - BIDIPHAR",
                  ps(14, bold=True, color="#c0392b", sa=4)),
        Paragraph(f"File: {file_name}  |  Xuất lúc: {now}", ps(8, sa=4)),
        Spacer(1, 0.2*cm),
    ]
    def h(t): story.append(Paragraph(t, ps(10, bold=True, color="#2c3e50", sb=8, sa=3)))

    h("CHỈ SỐ ĐẶT GIAO")
    story.append(mktbl([
        ["CHỈ SỐ ĐẶT GIAO",               "Giá trị"],
        ["Tổng Giá Trị Đặt",      f"{total_val:,.2f} Tr.đồng"],
        ["Đã Giao Hàng",          f"{delivered_val:,.2f} Tr.đồng"],
        ["Chưa Giao (Còn lại)",   f"{remain_val:,.2f} Tr.đồng"],
        ["Tỷ lệ Chưa Giao (DT)",  f"{rate_v:.1f}%"],
        ["Tổng Đơn Đặt (SQ/SO)",  f"{t_ord:,} Đơn"],
        ["Số Đơn Đã Xuất",        f"{d_ord:,} Đơn"],
        ["Số Đơn Chưa Xuất",      f"{r_ord:,} Đơn"],
        ["Tỷ lệ Chưa Xuất (Đơn)", f"{rate_o:.1f}%"],
    ], [10*cm, 6*cm], "#c0392b"))

    if not df_res.empty:
        h("NGUYÊN NHÂN CHƯA XUẤT")
        rows = [["Lý Do","Giá trị (Tr.đồng)","Số đơn"]]
        for _, r in df_res.sort_values("Trd",ascending=False).iterrows():
            rows.append([str(r["Ly Do"]),f"{r['Trd']:,.2f}",str(int(r["So don"]))])
        story.append(mktbl(rows,[10*cm,4*cm,3*cm],"#e67e22"))

    if not df_asm.empty:
        h("TOP 10 ASM CÓ DOANH SỐ CHƯA XUẤT CAO NHẤT")
        rows = [["Tên ASM","Chưa xuất (Tr.đồng)"]]
        for _, r in df_asm.iterrows():
            rows.append([str(r["Ten ASM"]),f"{r['Trd']:,.2f}"])
        story.append(mktbl(rows,[12*cm,5*cm],"#c0392b"))

    if not df_p.empty:
        h("TOP 10 SẢN PHẨM CÓ DOANH SỐ CHƯA XUẤT CAO NHẤT")
        rows = [["Tên Sản Phẩm","Chưa xuất (Tr.đồng)"]]
        for _, r in df_p.iterrows():
            rows.append([str(r["Ten SP"]),f"{r['Trd']:,.2f}"])
        story.append(mktbl(rows,[12*cm,5*cm],"#e67e22"))

    if not df_detail.empty:
        h("BẢNG CHI TIẾT: TOP 5 ASM × TOP 3 SẢN PHẨM CHƯA XUẤT")
        rows = [["Quản lý (ASM)","Sản Phẩm Chưa Xuất","Giá Trị Chưa Xuất"]]
        for _, r in df_detail.iterrows():
            rows.append([str(r.iloc[0]),str(r.iloc[1]),str(r.iloc[2])])
        story.append(mktbl(rows,[5.5*cm,7.5*cm,4*cm],"#2c3e50"))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# =============================================================================
# TIÊU ĐỀ + NÚT PDF ĐẦU TRANG
# Dùng st.rerun(): lần 1 tính PDF → lưu session_state → rerun
#                  lần 2 trở đi: PDF đã có → nút active ngay
# =============================================================================
import datetime as _dt

_ct, _cb = st.columns([8, 2])
with _ct:
    st.markdown("# 📋 PHÂN TÍCH HIỆU SUẤT GIAO HÀNG DMS")
with _cb:
    st.markdown("<div style='padding-top:14px'>", unsafe_allow_html=True)
    if "pdf_bytes" in st.session_state and st.session_state["pdf_bytes"]:
        st.download_button(
            label="📥 Xuất báo cáo PDF",
            data=st.session_state["pdf_bytes"],
            file_name=f"BaoCao_DMS_{_dt.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
    else:
        st.button("📥 Xuất báo cáo PDF", disabled=True,
                  use_container_width=True, help="Đang chuẩn bị...")
    st.markdown("</div>", unsafe_allow_html=True)
st.markdown("---")



# -----------------------------------------------------------------------------
# 1. ĐỌC FILE EXCEL MỚI NHẤT TRONG CÙNG THƯ MỤC (DÙNG KHI DEPLOY LÊN CLOUD)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=300)
def load_and_clean_data():
    # Tìm tất cả file Excel trong cùng thư mục với app
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_pattern = os.path.join(base_dir, "*.xls*")
    list_of_files = glob.glob(file_pattern)

    if not list_of_files:
        return None

    latest_file = max(list_of_files, key=os.path.getmtime)
    df = pd.read_excel(latest_file)

    df = df.dropna(how='all')
    df.columns = df.columns.astype(str).str.strip()

    if 'Lý Do' in df.columns:
        df['Lý Do'] = df['Lý Do'].fillna('Chưa có nguyên nhân').astype(str).str.strip()
        df.loc[df['Lý Do'] == '', 'Lý Do'] = 'Chưa có nguyên nhân'
    else:
        df['Lý Do'] = 'Chưa có nguyên nhân'

    num_cols = ['Giá Trị Đặt Hàng', 'Giá Trị Giao Hàng', 'Giá Trị Còn Lại',
                'SL đặt theo ĐVT CB', 'SL giao theo ĐVT CB', 'SL còn lại theo ĐVT CB']
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df, os.path.basename(latest_file)

data_load = load_and_clean_data()
if data_load is None:
    st.error("❌ Không tìm thấy file báo cáo Excel nào trong thư mục!")
    st.stop()

df_raw, file_name = data_load

info_col1, info_col2 = st.columns([4, 1])
info_col1.info(f"📌 **File báo cáo DMS đang đọc:** {file_name}")
if info_col2.button("↺ Cập nhật dữ liệu (Refresh)", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

# Hiển thị khoảng thời gian đơn hàng
ngay_col = next((c for c in df_raw.columns if 'ngày' in c.lower() and 'đơn' in c.lower()), None)
if ngay_col is None:
    ngay_col = next((c for c in df_raw.columns if 'ngay' in c.lower() or 'date' in c.lower()), None)
if ngay_col:
    ngay_series = pd.to_datetime(df_raw[ngay_col], errors='coerce').dropna()
    if not ngay_series.empty:
        ngay_min = ngay_series.min().strftime('%d/%m/%Y')
        ngay_max = ngay_series.max().strftime('%d/%m/%Y')
        st.caption(f"📆 **Thời gian đơn hàng:** từ **{ngay_min}** đến **{ngay_max}**")

# -----------------------------------------------------------------------------
# 2. BỘ LỌC NGANG ĐẦU TRANG — CASCADE (liên kết nhau)
# -----------------------------------------------------------------------------

# CSS làm nổi tiêu đề bộ lọc
st.markdown("""
<style>
div[data-testid="stSelectbox"] label,
div[data-testid="stMultiSelect"] label {
    font-weight: 800 !important;
    color: #1a1a1a !important;
    font-size: 0.95rem !important;
}
div[data-testid="column"]:nth-child(1) div[data-testid="stSelectbox"] > div:first-child,
div[data-testid="column"]:nth-child(1) div[data-testid="stMultiSelect"] > div:first-child {
    background-color: #fff3e0;
    border-radius: 8px;
    padding: 4px 8px;
}
div[data-testid="column"]:nth-child(2) div[data-testid="stMultiSelect"] > div:first-child {
    background-color: #e8f5e9;
    border-radius: 8px;
    padding: 4px 8px;
}
div[data-testid="column"]:nth-child(3) div[data-testid="stMultiSelect"] > div:first-child {
    background-color: #e3f2fd;
    border-radius: 8px;
    padding: 4px 8px;
}
div[data-testid="column"]:nth-child(4) div[data-testid="stMultiSelect"] > div:first-child {
    background-color: #f3e5f5;
    border-radius: 8px;
    padding: 4px 8px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("### 📌 Bộ lọc dữ liệu")
f1, f2, f3, f4 = st.columns(4)

def get_col_opts(df, col):
    if col in df.columns:
        return sorted(df[col].dropna().unique().tolist())
    return []

# --- BƯỚC 1: Chọn Kênh ---
with f1:
    s_kenh = st.selectbox("🟠 Kênh", ["Tất cả"] + get_col_opts(df_raw, 'Kênh'))

# Lọc theo Kênh trước để cascade xuống
df_after_kenh = df_raw.copy()
if s_kenh != "Tất cả" and 'Kênh' in df_after_kenh.columns:
    df_after_kenh = df_after_kenh[df_after_kenh['Kênh'] == s_kenh]

# --- BƯỚC 2: Chọn ASM (chỉ hiện ASM thuộc Kênh đã chọn) ---
with f2:
    s_asm = st.multiselect(
        "🟢 Quản lý (ASM)",
        get_col_opts(df_after_kenh, 'Tên ASM'),
        placeholder="Chọn hoặc gõ tên ASM..."
    )

# Lọc tiếp theo ASM
df_after_asm = df_after_kenh.copy()
if s_asm and 'Tên ASM' in df_after_asm.columns:
    df_after_asm = df_after_asm[df_after_asm['Tên ASM'].isin(s_asm)]

# --- BƯỚC 3: Chọn NVBH (chỉ hiện NVBH thuộc Kênh + ASM đã chọn) ---
with f3:
    s_nvbh = st.multiselect(
        "🔵 Nhân viên bán hàng",
        get_col_opts(df_after_asm, 'Tên NVBH'),
        placeholder="Chọn hoặc gõ tên NVBH..."
    )

# Lọc tiếp theo NVBH
df_after_nvbh = df_after_asm.copy()
if s_nvbh and 'Tên NVBH' in df_after_nvbh.columns:
    df_after_nvbh = df_after_nvbh[df_after_nvbh['Tên NVBH'].isin(s_nvbh)]

# --- BƯỚC 4: Chọn Sản phẩm (chỉ hiện SP thuộc các bộ lọc trên) ---
with f4:
    s_sp = st.multiselect(
        "🟣 Sản phẩm",
        get_col_opts(df_after_nvbh, 'Tên Sản Phẩm'),
        placeholder="Chọn hoặc gõ tên sản phẩm..."
    )

# DataFrame cuối cùng sau tất cả bộ lọc
df_f = df_after_nvbh.copy()
if s_sp and 'Tên Sản Phẩm' in df_f.columns:
    df_f = df_f[df_f['Tên Sản Phẩm'].isin(s_sp)]

# -----------------------------------------------------------------------------
# 3. LOGIC TÍNH TOÁN KPI
# -----------------------------------------------------------------------------
to_million = 1e6

total_val = df_f['Giá Trị Đặt Hàng'].sum() / to_million if 'Giá Trị Đặt Hàng' in df_f.columns else 0
delivered_val = df_f['Giá Trị Giao Hàng'].sum() / to_million if 'Giá Trị Giao Hàng' in df_f.columns else 0
remain_val = df_f['Giá Trị Còn Lại'].sum() / to_million if 'Giá Trị Còn Lại' in df_f.columns else 0
rate_v = (remain_val / total_val * 100) if total_val > 0 else 0

order_col = 'SQ/SO Order Code'
if order_col in df_f.columns and 'Giá Trị Còn Lại' in df_f.columns:
    df_ord = df_f.groupby(order_col).agg({'Giá Trị Còn Lại': 'sum'}).reset_index()
    t_ord = df_ord[order_col].nunique()
    r_ord = len(df_ord[df_ord['Giá Trị Còn Lại'] > 0.01])
    d_ord = t_ord - r_ord
    rate_o = (r_ord / t_ord * 100) if t_ord > 0 else 0
else:
    t_ord, d_ord, r_ord, rate_o = 0, 0, 0, 0

st.markdown("### 📊 Chỉ số KPI hiệu suất")
r1 = st.columns(4)
r1[0].metric("📊 Tổng Giá Trị Đặt", f"{total_val:,.2f} Trđ")
r1[1].metric("📦 Đã Giao Hàng", f"{delivered_val:,.2f} Trđ")
r1[2].metric("📌 Chưa Giao (Còn lại)", f"{remain_val:,.2f} Trđ")
r1[3].metric("📉 Tỷ lệ Chưa Giao (Doanh thu)", f"{rate_v:.1f}%")

r2 = st.columns(4)
r2[0].metric("📋 Tổng Số Đơn Đặt (SQ/SO)", f"{t_ord:,.0f} Đơn")
r2[1].metric("✔️ Số Đơn Đã Xuất", f"{d_ord:,.0f} Đơn")
r2[2].metric("🔴 Số Đơn Chưa Xuất", f"{r_ord:,.0f} Đơn")
r2[3].metric("📉 Tỷ lệ Chưa Xuất (Đơn hàng)", f"{rate_o:.1f}%")

st.markdown("---")

# -----------------------------------------------------------------------------
# 4. BIỂU ĐỒ BÁNH SONG SONG
# -----------------------------------------------------------------------------
c1, c2 = st.columns(2)

with c1:
    st.subheader("📊 Tỷ lệ Hoàn thành theo Doanh thu")
    fig_v = go.Figure(data=[go.Pie(
        labels=['ĐÃ GIAO (Trđ)', 'CHƯA GIAO (Trđ)'],
        values=[delivered_val, remain_val],
        hole=.4,
        marker_colors=['#2ecc71', '#e74c3c'],
        textinfo='label+value+percent',
        insidetextorientation='horizontal'
    )])
    fig_v.update_traces(
        texttemplate="<b>%{label}</b><br><b>%{value:,.2f} Trđ</b><br><b>%{percent}</b>",
        textfont=dict(size=13, color='white'),
        textposition='inside'
    )
    fig_v.update_layout(height=400, margin=dict(t=30, b=0, l=0, r=0))
    st.plotly_chart(fig_v, use_container_width=True)

with c2:
    st.subheader("📋 Tỷ lệ Hoàn thành theo Đơn hàng")
    fig_o = go.Figure(data=[go.Pie(
        labels=['ĐƠN ĐÃ XUẤT', 'ĐƠN CHƯA XUẤT'],
        values=[d_ord, r_ord],
        hole=.4,
        marker_colors=['#3498db', '#f39c12'],
        textinfo='label+value+percent',
        insidetextorientation='horizontal'
    )])
    fig_o.update_traces(
        texttemplate="<b>%{label}</b><br><b>%{value:,.0f} Đơn</b><br><b>%{percent}</b>",
        textfont=dict(size=13, color='white'),
        textposition='inside'
    )
    fig_o.update_layout(height=400, margin=dict(t=30, b=0, l=0, r=0))
    st.plotly_chart(fig_o, use_container_width=True)

# -----------------------------------------------------------------------------
# 5. NGUYÊN NHÂN & TOP ASM
# -----------------------------------------------------------------------------
st.markdown("---")
c3, c4 = st.columns(2)

with c3:
    st.subheader("📌 Nguyên nhân chưa xuất")
    if 'Giá Trị Còn Lại' in df_f.columns:
        # Tính doanh số và số đơn chưa xuất theo lý do
        df_res_val = df_f.groupby('Lý Do')['Giá Trị Còn Lại'].sum().reset_index()
        df_res_val['Trđ'] = df_res_val['Giá Trị Còn Lại'] / to_million

        # Đếm số đơn chưa xuất (Giá Trị Còn Lại > 0) theo lý do
        df_res_don = df_f[df_f['Giá Trị Còn Lại'] > 0.01].groupby('Lý Do').size().reset_index(name='Số đơn')

        df_res = df_res_val.merge(df_res_don, on='Lý Do', how='left')
        df_res['Số đơn'] = df_res['Số đơn'].fillna(0).astype(int)
        df_res = df_res.sort_values('Trđ', ascending=True)

        # Tạo nhãn kết hợp doanh số + số đơn
        df_res['label'] = df_res.apply(
            lambda r: f"{r['Trđ']:,.2f} Trđ  |  {r['Số đơn']:,} đơn", axis=1
        )

        fig_res = px.bar(df_res, x='Trđ', y='Lý Do', orientation='h', text='label',
                         color='Trđ', color_continuous_scale='Oranges', labels={'Trđ': 'Giá trị chưa xuất (Trđ)'})
        fig_res.update_traces(
            textposition='auto',
            insidetextanchor='middle',
            insidetextfont=dict(size=12, color='white'),
            outsidetextfont=dict(size=12, color='black')
        )
        fig_res.update_yaxes(tickfont=dict(size=12, color='black'))
        fig_res.update_xaxes(automargin=True)
        fig_res.update_layout(
            height=420,
            coloraxis_showscale=False,
            margin=dict(t=10, b=10, l=10, r=220),
            uniformtext_minsize=11,
            uniformtext_mode='show'
        )
        st.plotly_chart(fig_res, use_container_width=True)
    else:
        st.warning("Thiếu dữ liệu cột Giá Trị Còn Lại")

with c4:
    st.subheader("📊 Top 10 ASM có doanh số chưa xuất cao nhất")
    if 'Tên ASM' in df_f.columns and 'Giá Trị Còn Lại' in df_f.columns:
        df_asm = df_f.groupby('Tên ASM')['Giá Trị Còn Lại'].sum().reset_index()
        df_asm['Trđ'] = df_asm['Giá Trị Còn Lại'] / to_million
        df_asm = df_asm.sort_values('Trđ', ascending=False).head(10)
        fig_asm = px.bar(df_asm, x='Tên ASM', y='Trđ', color='Trđ', color_continuous_scale='Reds',
                         labels={'Trđ': 'Chưa xuất (Trđ)'}, text='Trđ')
        fig_asm.update_traces(
            texttemplate='<b>%{text:,.2f}</b>',
            textposition='outside',
            textfont=dict(size=12, color='black')
        )
        fig_asm.update_xaxes(tickfont=dict(size=11, color='black'), tickangle=-30)
        fig_asm.update_layout(
            height=420,
            coloraxis_showscale=False,
            margin=dict(t=40, b=10, l=10, r=10),
            uniformtext_minsize=10,
            uniformtext_mode='show'
        )
        st.plotly_chart(fig_asm, use_container_width=True)

# -----------------------------------------------------------------------------
# 6. TOP 10 SẢN PHẨM
# -----------------------------------------------------------------------------
st.markdown("---")
st.subheader("📊 Top 10 sản phẩm có doanh số chưa xuất cao nhất")
if 'Tên Sản Phẩm' in df_f.columns and 'Giá Trị Còn Lại' in df_f.columns:
    df_p = df_f.groupby('Tên Sản Phẩm')['Giá Trị Còn Lại'].sum().reset_index()
    df_p['Trđ'] = df_p['Giá Trị Còn Lại'] / to_million
    df_p = df_p.sort_values('Trđ', ascending=False).head(10)
    fig_p = px.bar(df_p, x='Tên Sản Phẩm', y='Trđ', color='Trđ', color_continuous_scale='YlOrRd',
                   labels={'Trđ': 'Chưa xuất (Trđ)'}, text='Trđ')
    fig_p.update_traces(
        texttemplate='<b>%{text:,.2f}</b>',
        textposition='outside',
        textfont=dict(size=12, color='black')
    )
    fig_p.update_xaxes(tickfont=dict(size=11, color='black'), tickangle=-30)
    fig_p.update_layout(
        height=460,
        coloraxis_showscale=False,
        margin=dict(t=40, b=10, l=10, r=10),
        uniformtext_minsize=10,
        uniformtext_mode='show'
    )
    st.plotly_chart(fig_p, use_container_width=True)

# -----------------------------------------------------------------------------
# 7. BẢNG TOP 5 ASM x TOP 3 SẢN PHẨM
# -----------------------------------------------------------------------------
st.markdown("---")
st.subheader("🔎 Bảng chi tiết: Top 5 ASM và Top 3 sản phẩm chưa xuất cao nhất theo từng ASM")

if 'Tên ASM' in df_f.columns and 'Tên Sản Phẩm' in df_f.columns and 'Giá Trị Còn Lại' in df_f.columns:
    df_asm_total = df_f.groupby('Tên ASM')['Giá Trị Còn Lại'].sum().reset_index()
    top5_asm = df_asm_total.sort_values('Giá Trị Còn Lại', ascending=False).head(5)['Tên ASM'].tolist()
    df_top5 = df_f[df_f['Tên ASM'].isin(top5_asm)].copy()
    df_dd = df_top5.groupby(['Tên ASM', 'Tên Sản Phẩm'])['Giá Trị Còn Lại'].sum().reset_index()
    df_dd['Doanh số chưa xuất (Trđ)'] = df_dd['Giá Trị Còn Lại'] / to_million
    df_dd['Rank'] = df_dd.groupby('Tên ASM')['Doanh số chưa xuất (Trđ)'].rank(method='first', ascending=False)
    df_top3 = df_dd[df_dd['Rank'] <= 3].copy()
    asm_total_map = df_asm_total.set_index('Tên ASM')['Giá Trị Còn Lại'] / to_million
    df_top3['Tổng đọng ASM'] = df_top3['Tên ASM'].map(asm_total_map)
    df_top3 = df_top3.sort_values(by=['Tổng đọng ASM', 'Doanh số chưa xuất (Trđ)'], ascending=[False, False])
    df_show = df_top3[['Tên ASM', 'Tên Sản Phẩm', 'Doanh số chưa xuất (Trđ)']].copy()
    df_show['Doanh số chưa xuất (Trđ)'] = df_show['Doanh số chưa xuất (Trđ)'].map('{:,.2f} Trđ'.format)
    df_show.columns = ['Quản lý (ASM)', 'Sản Phẩm Chưa Xuất', 'Giá Trị Chưa Xuất']
    df_show = df_show.reset_index(drop=True)
    st.dataframe(df_show, use_container_width=True, hide_index=True)
else:
    st.warning("Không thể hiển thị bảng chi tiết do thiếu cột dữ liệu cần thiết.")







# =============================================================================
# TÍNH DATA PDF + LƯU SESSION_STATE + RERUN (chỉ lần đầu)
# =============================================================================
if "pdf_bytes" not in st.session_state or st.session_state.get("pdf_filter") != str(df_f.shape):

    _df_res_pdf = pd.DataFrame()
    if "Lý Do" in df_f.columns and "Giá Trị Còn Lại" in df_f.columns:
        _df_res_pdf = df_f.groupby("Lý Do")["Giá Trị Còn Lại"].sum().reset_index()
        _df_res_pdf["Trd"]   = _df_res_pdf["Giá Trị Còn Lại"] / to_million
        _df_res_pdf["Ly Do"] = _df_res_pdf["Lý Do"]
        _df_don = (df_f[df_f["Giá Trị Còn Lại"] > 0.01]
                   .groupby("Lý Do").size().reset_index(name="So don"))
        _df_res_pdf = _df_res_pdf.merge(_df_don, on="Lý Do", how="left")
        _df_res_pdf["So don"] = _df_res_pdf["So don"].fillna(0).astype(int)

    _df_asm_pdf = pd.DataFrame()
    if "Tên ASM" in df_f.columns and "Giá Trị Còn Lại" in df_f.columns:
        _df_asm_pdf = df_f.groupby("Tên ASM")["Giá Trị Còn Lại"].sum().reset_index()
        _df_asm_pdf["Trd"]     = _df_asm_pdf["Giá Trị Còn Lại"] / to_million
        _df_asm_pdf["Ten ASM"] = _df_asm_pdf["Tên ASM"]
        _df_asm_pdf = _df_asm_pdf.sort_values("Trd", ascending=False).head(10)

    _df_p_pdf = pd.DataFrame()
    if "Tên Sản Phẩm" in df_f.columns and "Giá Trị Còn Lại" in df_f.columns:
        _df_p_pdf = df_f.groupby("Tên Sản Phẩm")["Giá Trị Còn Lại"].sum().reset_index()
        _df_p_pdf["Trd"]    = _df_p_pdf["Giá Trị Còn Lại"] / to_million
        _df_p_pdf["Ten SP"] = _df_p_pdf["Tên Sản Phẩm"]
        _df_p_pdf = _df_p_pdf.sort_values("Trd", ascending=False).head(10)

    _df_detail_pdf = df_show if "df_show" in dir() else pd.DataFrame()

    try:
        st.session_state["pdf_bytes"] = _build_pdf(
            file_name, total_val, delivered_val, remain_val, rate_v,
            t_ord, d_ord, r_ord, rate_o,
            _df_res_pdf, _df_asm_pdf, _df_p_pdf, _df_detail_pdf
        )
        st.session_state["pdf_filter"] = str(df_f.shape)
    except Exception as _e:
        st.session_state["pdf_bytes"]  = None
        st.session_state["pdf_filter"] = ""
        st.error(f"Loi tao PDF: {_e}")

    st.rerun()
