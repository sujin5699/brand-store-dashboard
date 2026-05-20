import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import re
import numpy as np
import time
import threading
import io
from datetime import datetime
from streamlit_autorefresh import st_autorefresh
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

DEFAULT_FOLDER = r"C:\Users\Sujin LEE\.claude\브랜드 스토어 상품성과 로우데이터"


@st.cache_resource
def _shared_state():
    """서버 생존 기간 동안 유지되는 전역 상태 (모든 세션 공유)"""
    return {"reload_needed": False, "known_files": frozenset()}


@st.cache_resource
def _start_file_watcher(folder: str):
    """11시 정각에 폴더를 스캔해 새 파일 감지 — 서버당 한 번만 시작"""
    state = _shared_state()

    def _run():
        alerted_this_hour = False
        while True:
            now = datetime.now()
            if now.hour == 11 and not alerted_this_hour:
                try:
                    current = frozenset(f for f in os.listdir(folder) if f.endswith(".xlsx"))
                    if current != state["known_files"]:
                        state["reload_needed"] = True
                    alerted_this_hour = True
                except Exception:
                    pass
            elif now.hour != 11:
                alerted_this_hour = False
            time.sleep(30)

    threading.Thread(target=_run, daemon=True).start()
    return True

st.set_page_config(
    page_title="브랜드 스토어 판매 대시보드",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .metric-card {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 16px;
        border-left: 4px solid #4CAF50;
    }
    .insight-card {
        background: #fff3cd;
        border-radius: 8px;
        padding: 12px;
        margin: 6px 0;
        border-left: 4px solid #ffc107;
    }
    .insight-danger {
        background: #f8d7da;
        border-left-color: #dc3545;
    }
    .insight-success {
        background: #d4edda;
        border-left-color: #28a745;
    }
    h1 { color: #1a1a2e; }
</style>
""", unsafe_allow_html=True)


def parse_date_from_filename(filename: str) -> str | None:
    m = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
    return m.group(1) if m else None


def load_excel_files(files) -> pd.DataFrame:
    dfs = []
    for f in files:
        date_str = parse_date_from_filename(f.name)
        if not date_str:
            st.warning(f"날짜를 파싱할 수 없습니다: {f.name}")
            continue
        try:
            df = pd.read_excel(f, engine="openpyxl")
            df["날짜"] = pd.to_datetime(date_str)
            df["파일명"] = f.name
            dfs.append(df)
        except Exception as e:
            st.error(f"파일 읽기 오류 {f.name}: {e}")
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)


@st.cache_data(ttl=3600, show_spinner="폴더에서 데이터 로딩 중...")
def load_from_folder(folder: str) -> pd.DataFrame:
    dfs = []
    for fname in os.listdir(folder):
        if not fname.endswith(".xlsx"):
            continue
        date_str = parse_date_from_filename(fname)
        if not date_str:
            continue
        try:
            df = pd.read_excel(os.path.join(folder, fname), engine="openpyxl")
            df["날짜"] = pd.to_datetime(date_str)
            df["파일명"] = fname
            dfs.append(df)
        except Exception as e:
            st.warning(f"파일 읽기 오류 {fname}: {e}")
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)


@st.cache_data(ttl=3600, show_spinner="Google Drive에서 데이터 로딩 중...")
def load_from_gdrive(folder_id: str, creds_info: dict) -> pd.DataFrame:
    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    results = service.files().list(
        q=(
            f"'{folder_id}' in parents"
            " and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
            " and trashed=false"
        ),
        fields="files(id, name)",
        orderBy="name",
    ).execute()

    dfs = []
    for f in results.get("files", []):
        date_str = parse_date_from_filename(f["name"])
        if not date_str:
            continue
        try:
            request = service.files().get_media(fileId=f["id"])
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            buf.seek(0)
            df = pd.read_excel(buf, engine="openpyxl")
            df["날짜"] = pd.to_datetime(date_str)
            df["파일명"] = f["name"]
            dfs.append(df)
        except Exception as e:
            st.warning(f"Drive 파일 읽기 오류 {f['name']}: {e}")

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def aggregate_daily(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.groupby("날짜")
        .agg(
            결제수=("결제수", "sum"),
            결제상품수량=("결제상품수량", "sum"),
            결제금액=("결제금액", "sum"),
            쿠폰합계=("쿠폰합계", "sum"),
            환불건수=("환불건수", "sum"),
            환불금액=("환불금액", "sum"),
            환불수량=("환불수량", "sum"),
            모바일비율=("모바일비율(결제금액)", "mean"),
        )
        .reset_index()
        .sort_values("날짜")
    )


def aggregate_weekly(daily: pd.DataFrame) -> pd.DataFrame:
    d = daily.copy()
    d["주차"] = d["날짜"].dt.to_period("W").dt.start_time
    return (
        d.groupby("주차")
        .agg(
            결제수=("결제수", "sum"),
            결제상품수량=("결제상품수량", "sum"),
            결제금액=("결제금액", "sum"),
            쿠폰합계=("쿠폰합계", "sum"),
            환불건수=("환불건수", "sum"),
            환불금액=("환불금액", "sum"),
            환불수량=("환불수량", "sum"),
            모바일비율=("모바일비율", "mean"),
        )
        .reset_index()
        .rename(columns={"주차": "날짜"})
    )


def aggregate_monthly(daily: pd.DataFrame) -> pd.DataFrame:
    d = daily.copy()
    d["월"] = d["날짜"].dt.to_period("M").dt.start_time
    return (
        d.groupby("월")
        .agg(
            결제수=("결제수", "sum"),
            결제상품수량=("결제상품수량", "sum"),
            결제금액=("결제금액", "sum"),
            쿠폰합계=("쿠폰합계", "sum"),
            환불건수=("환불건수", "sum"),
            환불금액=("환불금액", "sum"),
            환불수량=("환불수량", "sum"),
            모바일비율=("모바일비율", "mean"),
        )
        .reset_index()
        .rename(columns={"월": "날짜"})
    )


def detect_anomalies(daily: pd.DataFrame, col: str, threshold: float = 1.5) -> pd.DataFrame:
    if len(daily) < 3:
        return pd.DataFrame()
    d = daily[["날짜", col]].copy()
    mean = d[col].mean()
    std = d[col].std()
    if std == 0:
        return pd.DataFrame()
    d["z_score"] = (d[col] - mean) / std
    d["pct_vs_avg"] = (d[col] - mean) / mean * 100
    return d[d["z_score"].abs() > threshold].sort_values("z_score")


def generate_insights(daily: pd.DataFrame, df_raw: pd.DataFrame) -> list[dict]:
    insights = []
    if len(daily) < 2:
        return insights

    latest = daily.iloc[-1]
    prev = daily.iloc[-2]

    def pct(a, b):
        return (a - b) / b * 100 if b != 0 else 0

    rev_chg = pct(latest["결제금액"], prev["결제금액"])
    if abs(rev_chg) >= 20:
        sign = "급등" if rev_chg > 0 else "급락"
        insights.append({
            "type": "success" if rev_chg > 0 else "danger",
            "title": f"결제금액 {sign}",
            "body": f"전일 대비 {rev_chg:+.1f}% ({int(prev['결제금액']):,}원 → {int(latest['결제금액']):,}원)",
        })

    refund_rate_latest = (latest["환불금액"] / latest["결제금액"] * 100) if latest["결제금액"] else 0
    refund_rate_prev = (prev["환불금액"] / prev["결제금액"] * 100) if prev["결제금액"] else 0
    if refund_rate_latest > 10:
        insights.append({
            "type": "danger",
            "title": "환불율 높음 주의",
            "body": f"최근일 환불율 {refund_rate_latest:.1f}% (전일 {refund_rate_prev:.1f}%)",
        })

    coupon_rate = (latest["쿠폰합계"] / latest["결제금액"] * 100) if latest["결제금액"] else 0
    if coupon_rate > 30:
        insights.append({
            "type": "warning",
            "title": "쿠폰 의존도 높음",
            "body": f"결제금액 대비 쿠폰 비율 {coupon_rate:.1f}% — 수익성 점검 필요",
        })

    mobile = latest["모바일비율"] * 100
    if mobile > 85:
        insights.append({
            "type": "success",
            "title": "모바일 비중 높음",
            "body": f"모바일 결제 비율 {mobile:.1f}% — 모바일 최적화 유지 권장",
        })

    anomalies = detect_anomalies(daily, "결제금액")
    for _, row in anomalies.iterrows():
        date_str = row["날짜"].strftime("%m/%d")
        insights.append({
            "type": "danger" if row["z_score"] < 0 else "warning",
            "title": f"{date_str} 결제금액 이상 감지 (Z={row['z_score']:.1f})",
            "body": f"평균 대비 {row['pct_vs_avg']:+.1f}% ({int(row['결제금액']):,}원)",
        })

    if len(df_raw) > 0 and "상품명" in df_raw.columns:
        top_products = (
            df_raw.groupby("상품명")["결제금액"].sum().nlargest(3).reset_index()
        )
        prod_list = ", ".join(
            f"{r['상품명'][:15]} ({int(r['결제금액']):,}원)"
            for _, r in top_products.iterrows()
        )
        insights.append({
            "type": "info",
            "title": "매출 상위 상품",
            "body": prod_list,
        })

    return insights


def fmt_won(v: float) -> str:
    if v >= 1_000_000:
        return f"{v / 1_000_000:.1f}M원"
    elif v >= 1_000:
        return f"{v / 1_000:.0f}K원"
    return f"{int(v):,}원"


def render_trend_chart(agg: pd.DataFrame, title: str, x_fmt: str):
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=("결제금액", "결제수 / 결제상품수량", "환불금액 & 환불율", "모바일 비율"),
        vertical_spacing=0.14,
        horizontal_spacing=0.10,
    )

    x = agg["날짜"]

    fig.add_trace(go.Bar(x=x, y=agg["결제금액"], name="결제금액",
                         marker_color="#4CAF50", showlegend=False), row=1, col=1)

    fig.add_trace(go.Scatter(x=x, y=agg["결제수"], name="결제수",
                              mode="lines+markers", line=dict(color="#2196F3")), row=1, col=2)
    fig.add_trace(go.Scatter(x=x, y=agg["결제상품수량"], name="결제상품수량",
                              mode="lines+markers", line=dict(color="#9C27B0", dash="dash")), row=1, col=2)

    refund_rate = (agg["환불금액"] / agg["결제금액"].replace(0, np.nan) * 100).fillna(0)
    fig.add_trace(go.Bar(x=x, y=agg["환불금액"], name="환불금액",
                         marker_color="#f44336", showlegend=False), row=2, col=1)
    fig.add_trace(go.Scatter(x=x, y=refund_rate, name="환불율(%)",
                              mode="lines+markers", line=dict(color="#FF9800"),
                              yaxis="y5"), row=2, col=1)

    fig.add_trace(go.Scatter(x=x, y=agg["모바일비율"] * 100, name="모바일비율(%)",
                              mode="lines+markers", fill="tozeroy",
                              line=dict(color="#00BCD4"), showlegend=False), row=2, col=2)

    fig.update_layout(
        title=title,
        height=520,
        legend=dict(orientation="h", y=-0.12),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=60, b=40),
    )
    fig.update_xaxes(tickformat=x_fmt)
    fig.update_yaxes(tickformat=",", row=1, col=1)
    fig.update_yaxes(tickformat=",", row=1, col=2)
    return fig


# ── 자동 새로고침 (10분마다 — 11시 갱신 반영용) ───────────────────────────
st_autorefresh(interval=10 * 60 * 1000, key="auto_refresh")

# ── 데이터 소스 결정 ──────────────────────────────────────────────────────
USE_GDRIVE = "gdrive" in st.secrets

# ── 11시 자동 갱신 처리 ──────────────────────────────────────────────────
_state = _shared_state()

if USE_GDRIVE:
    # 클라우드: 10분 자동갱신 사이클 안에서 11시 감지
    now = datetime.now()
    last_date = _state.get("last_gdrive_date")
    if now.hour == 11 and last_date != now.date():
        _state["reload_needed"] = True
        _state["last_gdrive_date"] = now.date()
else:
    # 로컬: 백그라운드 스레드가 처리
    _start_file_watcher(DEFAULT_FOLDER)

if _state.get("reload_needed"):
    _state["reload_needed"] = False
    st.session_state.pop("df_raw", None)
    st.toast("새 파일이 감지되어 대시보드를 갱신합니다!", icon="🔄")

# ── Sidebar ──────────────────────────────────────────────────────────────────
st.sidebar.title("🛒 브랜드 스토어\n판매 대시보드")
st.sidebar.markdown("---")

df_raw = pd.DataFrame()

if USE_GDRIVE:
    # ── Google Drive 모드 (Streamlit Cloud) ──────────────────────────────
    folder_id = st.secrets["gdrive"]["folder_id"]
    creds_info = dict(st.secrets["gdrive"]["credentials"])

    if "df_raw" not in st.session_state:
        with st.spinner("Google Drive에서 데이터 로딩 중..."):
            df_raw = load_from_gdrive(folder_id, creds_info)
            st.session_state["df_raw"] = df_raw
    else:
        df_raw = st.session_state["df_raw"]

    if st.sidebar.button("수동 새로고침", use_container_width=True):
        with st.spinner("Google Drive에서 최신 데이터 로딩 중..."):
            df_raw = load_from_gdrive(folder_id, creds_info)
            st.session_state["df_raw"] = df_raw
        st.toast("갱신 완료!")

    st.sidebar.success("☁️ Google Drive 연동 중")
    st.sidebar.caption("매일 오전 11시 자동 갱신")

else:
    # ── 로컬 폴더 모드 ────────────────────────────────────────────────────
    source_mode = st.sidebar.radio("데이터 소스", ["📁 폴더 자동 로드", "📤 파일 업로드"])

    if source_mode == "📁 폴더 자동 로드":
        folder = st.sidebar.text_input("폴더 경로", value=DEFAULT_FOLDER)

        if "df_raw" not in st.session_state:
            with st.spinner("데이터 자동 로딩 중..."):
                df_raw = load_from_folder(folder)
                st.session_state["df_raw"] = df_raw
                _state["known_files"] = frozenset(f for f in os.listdir(folder) if f.endswith(".xlsx"))
        else:
            df_raw = st.session_state["df_raw"]

        if st.sidebar.button("수동 새로고침", use_container_width=True):
            with st.spinner("새로고침 중..."):
                df_raw = load_from_folder(folder)
                st.session_state["df_raw"] = df_raw
                _state["known_files"] = frozenset(f for f in os.listdir(folder) if f.endswith(".xlsx"))
            st.toast("데이터가 갱신되었습니다!")

        loaded_files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
        st.sidebar.caption(
            f"파일 {len(loaded_files)}개 로드됨  \n"
            f"마지막 확인: {datetime.now().strftime('%H:%M:%S')}"
        )
        st.sidebar.info("매일 오전 11시에 새 파일을 자동 감지합니다.")
    else:
        uploaded = st.sidebar.file_uploader(
            "엑셀 파일 선택 (복수 가능)",
            type=["xlsx"],
            accept_multiple_files=True,
        )
        if uploaded:
            with st.spinner("파일 로딩 중..."):
                df_raw = load_excel_files(uploaded)
                st.session_state["df_raw"] = df_raw
        elif "df_raw" in st.session_state:
            df_raw = st.session_state["df_raw"]

# ── Main ──────────────────────────────────────────────────────────────────────
st.title("🛒 브랜드 스토어 판매 성과 대시보드")

if df_raw.empty:
    st.info("데이터를 불러오는 중이거나, 폴더에 xlsx 파일이 없습니다.")
    st.stop()

# ── Filters ──────────────────────────────────────────────────────────────────
st.sidebar.markdown("---")
st.sidebar.subheader("필터")

if "상품카테고리(대)" in df_raw.columns:
    cats = ["전체"] + sorted(df_raw["상품카테고리(대)"].dropna().unique().tolist())
    sel_cat = st.sidebar.selectbox("상품 카테고리(대)", cats)
    if sel_cat != "전체":
        df_raw = df_raw[df_raw["상품카테고리(대)"] == sel_cat]

date_range = st.sidebar.date_input(
    "기간 선택",
    value=(df_raw["날짜"].min().date(), df_raw["날짜"].max().date()),
)
if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    df_raw = df_raw[
        (df_raw["날짜"].dt.date >= date_range[0]) &
        (df_raw["날짜"].dt.date <= date_range[1])
    ]

# ── Aggregations ─────────────────────────────────────────────────────────────
daily = aggregate_daily(df_raw)
weekly = aggregate_weekly(daily)
monthly = aggregate_monthly(daily)

# ── KPI Cards ────────────────────────────────────────────────────────────────
total_rev = daily["결제금액"].sum()
total_orders = daily["결제수"].sum()
total_refund = daily["환불금액"].sum()
avg_mobile = daily["모바일비율"].mean() * 100
refund_rate_total = total_refund / total_rev * 100 if total_rev else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("총 결제금액", fmt_won(total_rev))
col2.metric("총 결제건수", f"{int(total_orders):,}건")
col3.metric("총 환불금액", fmt_won(total_refund))
col4.metric("환불율", f"{refund_rate_total:.1f}%")
col5.metric("평균 모바일 비율", f"{avg_mobile:.1f}%")

if len(daily) >= 2:
    last = daily.iloc[-1]
    prev2 = daily.iloc[-2]
    pct_rev = (last["결제금액"] - prev2["결제금액"]) / prev2["결제금액"] * 100 if prev2["결제금액"] else 0
    pct_ord = (last["결제수"] - prev2["결제수"]) / prev2["결제수"] * 100 if prev2["결제수"] else 0
    st.caption(
        f"**최근일({last['날짜'].strftime('%m/%d')}) vs 전일**: "
        f"결제금액 {pct_rev:+.1f}%  |  결제건수 {pct_ord:+.1f}%"
    )

st.markdown("---")

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_prod_trend, tab_day, tab_week, tab_month, tab_product, tab_insight = st.tabs(
    ["📈 상품 판매 추이", "📅 일간", "📆 주간", "🗓️ 월간", "📦 상품 분석", "💡 인사이트"]
)

with tab_day:
    st.plotly_chart(render_trend_chart(daily, "일간 판매 추이", "%m/%d"), use_container_width=True)

    st.subheader("일별 상세 데이터")
    disp = daily.copy()
    disp["날짜"] = disp["날짜"].dt.strftime("%Y-%m-%d")
    disp["결제금액"] = disp["결제금액"].apply(lambda x: f"{int(x):,}")
    disp["환불금액"] = disp["환불금액"].apply(lambda x: f"{int(x):,}")
    disp["모바일비율"] = (disp["모바일비율"] * 100).map("{:.1f}%".format)
    st.dataframe(disp, use_container_width=True, hide_index=True)

with tab_week:
    if len(weekly) < 2:
        st.info("주간 집계를 위해 2주 이상 데이터가 필요합니다.")
    else:
        st.plotly_chart(render_trend_chart(weekly, "주간 판매 추이", "%m/%d 주"), use_container_width=True)
        disp = weekly.copy()
        disp["날짜"] = disp["날짜"].dt.strftime("%Y-%m-%d 주")
        disp["결제금액"] = disp["결제금액"].apply(lambda x: f"{int(x):,}")
        disp["환불금액"] = disp["환불금액"].apply(lambda x: f"{int(x):,}")
        disp["모바일비율"] = (disp["모바일비율"] * 100).map("{:.1f}%".format)
        st.dataframe(disp, use_container_width=True, hide_index=True)

with tab_month:
    if len(monthly) < 1:
        st.info("월간 집계 데이터가 없습니다.")
    else:
        st.plotly_chart(render_trend_chart(monthly, "월간 판매 추이", "%Y-%m"), use_container_width=True)
        disp = monthly.copy()
        disp["날짜"] = disp["날짜"].dt.strftime("%Y-%m")
        disp["결제금액"] = disp["결제금액"].apply(lambda x: f"{int(x):,}")
        disp["환불금액"] = disp["환불금액"].apply(lambda x: f"{int(x):,}")
        disp["모바일비율"] = (disp["모바일비율"] * 100).map("{:.1f}%".format)
        st.dataframe(disp, use_container_width=True, hide_index=True)

with tab_product:
    st.subheader("상품별 판매 성과")
    if "상품명" not in df_raw.columns:
        st.warning("상품명 컬럼이 없습니다.")
    else:
        prod_agg = (
            df_raw.groupby("상품명")
            .agg(결제금액=("결제금액", "sum"), 결제수=("결제수", "sum"),
                 환불금액=("환불금액", "sum"), 결제상품수량=("결제상품수량", "sum"))
            .reset_index()
            .sort_values("결제금액", ascending=False)
        )
        prod_agg["환불율"] = (prod_agg["환불금액"] / prod_agg["결제금액"].replace(0, np.nan) * 100).fillna(0)

        top_n = st.slider("상위 N개 상품", 5, min(30, len(prod_agg)), 10)
        top = prod_agg.head(top_n)

        col_a, col_b = st.columns(2)
        with col_a:
            fig_bar = px.bar(
                top, x="결제금액", y="상품명", orientation="h",
                title="상위 상품 결제금액", color="결제금액",
                color_continuous_scale="Greens",
                labels={"결제금액": "결제금액(원)", "상품명": ""},
            )
            fig_bar.update_layout(yaxis=dict(autorange="reversed"), height=400)
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_b:
            fig_scatter = px.scatter(
                prod_agg.head(top_n * 2), x="결제수", y="환불율",
                size="결제금액", hover_name="상품명",
                title="결제수 vs 환불율",
                labels={"결제수": "결제건수", "환불율": "환불율(%)"},
                color="환불율", color_continuous_scale="RdYlGn_r",
            )
            fig_scatter.update_layout(height=400)
            st.plotly_chart(fig_scatter, use_container_width=True)

        if "상품카테고리(대)" in df_raw.columns:
            cat_agg = (
                df_raw.groupby("상품카테고리(대)")
                .agg(결제금액=("결제금액", "sum"))
                .reset_index()
            )
            fig_pie = px.pie(cat_agg, names="상품카테고리(대)", values="결제금액",
                             title="카테고리별 매출 비중")
            st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader("전체 상품 상세")
        disp_prod = prod_agg.copy()
        disp_prod["결제금액"] = disp_prod["결제금액"].apply(lambda x: f"{int(x):,}")
        disp_prod["환불금액"] = disp_prod["환불금액"].apply(lambda x: f"{int(x):,}")
        disp_prod["환불율"] = disp_prod["환불율"].map("{:.1f}%".format)
        st.dataframe(disp_prod, use_container_width=True, hide_index=True)

with tab_prod_trend:
    st.subheader("상품별 판매 변화 추이")
    if "상품명" not in df_raw.columns:
        st.warning("상품명 컬럼이 없습니다.")
    else:
        n_unique = df_raw["상품명"].nunique()

        ctrl1, ctrl2, ctrl3 = st.columns([1, 1, 2])
        with ctrl1:
            top_n_trend = st.slider("상위 N개 상품", 3, min(20, n_unique), min(5, n_unique), key="pt_n")
        with ctrl2:
            metric_sel = st.selectbox(
                "지표",
                ["결제금액", "결제수", "결제상품수량", "환불금액"],
                key="pt_metric",
            )
        with ctrl3:
            period_sel = st.radio("기간 단위", ["일간", "주간", "월간"], horizontal=True, key="pt_period")

        # 상위 N 상품 (전체 기간 합산 기준)
        top_products = (
            df_raw.groupby("상품명")[metric_sel]
            .sum()
            .nlargest(top_n_trend)
            .index.tolist()
        )
        df_top = df_raw[df_raw["상품명"].isin(top_products)].copy()

        if period_sel == "일간":
            df_grouped = (
                df_top.groupby(["날짜", "상품명"])[metric_sel]
                .sum()
                .reset_index()
            )
            x_fmt = "%m/%d"
        elif period_sel == "주간":
            df_top["_기간"] = df_top["날짜"].dt.to_period("W").dt.start_time
            df_grouped = (
                df_top.groupby(["_기간", "상품명"])[metric_sel]
                .sum()
                .reset_index()
                .rename(columns={"_기간": "날짜"})
            )
            x_fmt = "%m/%d 주"
        else:
            df_top["_기간"] = df_top["날짜"].dt.to_period("M").dt.start_time
            df_grouped = (
                df_top.groupby(["_기간", "상품명"])[metric_sel]
                .sum()
                .reset_index()
                .rename(columns={"_기간": "날짜"})
            )
            x_fmt = "%Y-%m"

        # 상품명을 15자로 줄여서 범례 가독성 확보
        df_grouped["상품명_단축"] = df_grouped["상품명"].str[:18]

        # ── 멀티라인 차트 ──────────────────────────────────────────────
        fig_line = px.line(
            df_grouped,
            x="날짜", y=metric_sel, color="상품명_단축",
            markers=True,
            title=f"상위 {top_n_trend}개 상품 {period_sel} {metric_sel} 추이",
            labels={"상품명_단축": "상품명", metric_sel: metric_sel},
        )
        fig_line.update_layout(
            height=460,
            legend=dict(orientation="h", y=-0.25, font=dict(size=11)),
            xaxis_tickformat=x_fmt,
            hovermode="x unified",
            plot_bgcolor="white",
        )
        fig_line.update_traces(line=dict(width=2.5))
        st.plotly_chart(fig_line, use_container_width=True)

        # ── 누적 바차트 ────────────────────────────────────────────────
        fig_bar = px.bar(
            df_grouped,
            x="날짜", y=metric_sel, color="상품명_단축",
            title=f"상위 {top_n_trend}개 상품 {period_sel} {metric_sel} 구성",
            labels={"상품명_단축": "상품명"},
            barmode="stack",
        )
        fig_bar.update_layout(
            height=380,
            legend=dict(orientation="h", y=-0.25, font=dict(size=11)),
            xaxis_tickformat=x_fmt,
            plot_bgcolor="white",
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # ── 히트맵 ────────────────────────────────────────────────────
        pivot = (
            df_grouped.pivot_table(
                index="상품명_단축", columns="날짜", values=metric_sel, aggfunc="sum"
            )
            .fillna(0)
        )
        pivot.columns = [c.strftime(x_fmt.replace(" 주", "")) if hasattr(c, "strftime") else str(c)
                         for c in pivot.columns]

        fig_heat = go.Figure(
            go.Heatmap(
                z=pivot.values,
                x=pivot.columns.tolist(),
                y=pivot.index.tolist(),
                colorscale="Greens",
                hoverongaps=False,
                hovertemplate="날짜: %{x}<br>상품: %{y}<br>값: %{z:,.0f}<extra></extra>",
            )
        )
        fig_heat.update_layout(
            title=f"상품 × 날짜 히트맵 ({metric_sel})",
            height=max(300, top_n_trend * 40 + 100),
            xaxis_title="날짜",
            yaxis_title="",
            margin=dict(l=180),
        )
        st.plotly_chart(fig_heat, use_container_width=True)

        # ── 순위 변동 테이블 ──────────────────────────────────────────
        st.subheader("기간별 순위 변동")
        pivot_rank = (
            df_grouped.pivot_table(
                index="상품명_단축", columns="날짜", values=metric_sel, aggfunc="sum"
            )
            .fillna(0)
        )
        rank_df = pivot_rank.rank(ascending=False, method="min").astype(int)
        rank_df.columns = [c.strftime(x_fmt.replace(" 주", "")) if hasattr(c, "strftime") else str(c)
                           for c in rank_df.columns]
        st.dataframe(rank_df, use_container_width=True)

with tab_insight:
    st.subheader("💡 자동 인사이트 & 이상 감지")

    insights = generate_insights(daily, df_raw)
    if not insights:
        st.success("특이 사항이 발견되지 않았습니다.")
    else:
        for ins in insights:
            t = ins["type"]
            if t == "danger":
                color = "#f8d7da"
                border = "#dc3545"
                icon = "🔴"
            elif t == "success":
                color = "#d4edda"
                border = "#28a745"
                icon = "🟢"
            elif t == "info":
                color = "#d1ecf1"
                border = "#17a2b8"
                icon = "🔵"
            else:
                color = "#fff3cd"
                border = "#ffc107"
                icon = "🟡"
            st.markdown(
                f"""<div style="background:{color};border-left:4px solid {border};
                border-radius:8px;padding:10px 14px;margin:6px 0;">
                <b>{icon} {ins['title']}</b><br>{ins['body']}</div>""",
                unsafe_allow_html=True,
            )

    st.markdown("---")
    st.subheader("📈 결제금액 이상 탐지 (Z-score 기준)")
    anom = detect_anomalies(daily, "결제금액", threshold=1.0)
    if anom.empty:
        st.info("이상 데이터 없음 (데이터가 더 쌓이면 정확도 높아집니다)")
    else:
        fig_anom = px.scatter(
            daily, x="날짜", y="결제금액",
            title="결제금액 추이 (빨간 점: 이상 감지)",
        )
        fig_anom.add_scatter(
            x=anom["날짜"], y=anom["결제금액"],
            mode="markers", marker=dict(color="red", size=12, symbol="x"),
            name="이상 감지",
        )
        fig_anom.add_hline(
            y=daily["결제금액"].mean(), line_dash="dash", line_color="orange",
            annotation_text="평균",
        )
        st.plotly_chart(fig_anom, use_container_width=True)

    st.subheader("📊 요일별 평균 성과")
    d = daily.copy()
    d["요일"] = d["날짜"].dt.day_name()
    d["요일번호"] = d["날짜"].dt.dayofweek
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    day_kr = ["월", "화", "수", "목", "금", "토", "일"]
    dow = d.groupby(["요일번호", "요일"])["결제금액"].mean().reset_index().sort_values("요일번호")
    dow["요일KR"] = [day_kr[i] for i in dow["요일번호"]]
    fig_dow = px.bar(dow, x="요일KR", y="결제금액", title="요일별 평균 결제금액",
                     color="결제금액", color_continuous_scale="Blues")
    st.plotly_chart(fig_dow, use_container_width=True)
