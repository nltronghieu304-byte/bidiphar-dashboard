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
st.title("📋 PHÂN TÍCH HIỆU SUẤT GIAO HÀNG DMS")
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
# 2. BỘ LỌC NGANG ĐẦU TRANG
# -----------------------------------------------------------------------------
st.markdown("### 📌 Bộ lọc dữ liệu")
f1, f2, f3, f4 = st.columns(4)

def get_opts(df, col):
    if col in df.columns:
        return sorted(df[col].dropna().unique().tolist())
    return []

with f1:
    s_kenh = st.selectbox("Kênh", ["Tất cả"] + get_opts(df_raw, 'Kênh'))
with f2:
    s_asm = st.multiselect("Quản lý (ASM) — gõ để tìm", get_opts(df_raw, 'Tên ASM'), placeholder="Chọn hoặc gõ tên ASM...")
with f3:
    s_nvbh = st.multiselect("Nhân viên bán hàng — gõ để tìm", get_opts(df_raw, 'Tên NVBH'), placeholder="Chọn hoặc gõ tên NVBH...")
with f4:
    s_sp = st.multiselect("Sản phẩm — gõ để tìm", get_opts(df_raw, 'Tên Sản Phẩm'), placeholder="Chọn hoặc gõ tên sản phẩm...")

df_f = df_raw.copy()
if s_kenh != "Tất cả" and 'Kênh' in df_f.columns:
    df_f = df_f[df_f['Kênh'] == s_kenh]
if s_asm and 'Tên ASM' in df_f.columns:
    df_f = df_f[df_f['Tên ASM'].isin(s_asm)]
if s_nvbh and 'Tên NVBH' in df_f.columns:
    df_f = df_f[df_f['Tên NVBH'].isin(s_nvbh)]
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