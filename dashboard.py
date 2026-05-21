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
from datetime import datetime, timedelta
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


def _build_drive_service(creds_info: dict):
    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


@st.cache_data(ttl=300, show_spinner=False)
def _list_drive_files(folder_id: str, creds_info: dict) -> list[dict]:
    """파일 목록만 가져옴 (메타데이터만, 빠름)"""
    service = _build_drive_service(creds_info)
    results = service.files().list(
        q=(
            f"'{folder_id}' in parents"
            " and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
            " and trashed=false"
        ),
        fields="files(id, name)",
        orderBy="name",
    ).execute()
    return results.get("files", [])


@st.cache_data(show_spinner=False)
def _download_one_file(file_id: str, file_name: str, creds_info: dict) -> pd.DataFrame | None:
    """파일 1개 다운로드 — file_id가 같으면 영구 캐시"""
    date_str = parse_date_from_filename(file_name)
    if not date_str:
        return None
    try:
        service = _build_drive_service(creds_info)
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        df = pd.read_excel(buf, engine="openpyxl")
        df["날짜"] = pd.to_datetime(date_str)
        df["파일명"] = file_name
        return df
    except Exception as e:
        st.warning(f"Drive 파일 읽기 오류 {file_name}: {e}")
        return None


def load_from_gdrive(folder_id: str, creds_info: dict) -> pd.DataFrame:
    """파일 목록 확인 후 새 파일만 다운로드 (기존 파일은 캐시 사용)"""
    files = _list_drive_files(folder_id, creds_info)
    dfs = []
    progress = st.progress(0, text="데이터 로딩 중...")
    for i, f in enumerate(files):
        df = _download_one_file(f["id"], f["name"], creds_info)
        if df is not None:
            dfs.append(df)
        progress.progress((i + 1) / max(len(files), 1), text=f"로딩 중... ({i+1}/{len(files)})")
    progress.empty()
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


def detect_product_anomalies(
    df_grouped: pd.DataFrame,
    metric_col: str,
    name_col: str = "상품명_단축",
    pct_threshold: float = 30.0,
    min_ratio: float = 0.1,
) -> list[dict]:
    """직전 기간 대비 급락/급등 상품 반환 (pct_threshold % 이상 변화)"""
    dates = sorted(df_grouped["날짜"].unique())
    if len(dates) < 2:
        return []
    cur_date, prev_date = dates[-1], dates[-2]
    cur  = df_grouped[df_grouped["날짜"] == cur_date].set_index(name_col)[metric_col]
    prev = df_grouped[df_grouped["날짜"] == prev_date].set_index(name_col)[metric_col]

    # 최소 규모 필터: 상품 평균이 전체 평균의 min_ratio 미만이면 제외
    prod_mean  = df_grouped.groupby(name_col)[metric_col].mean()
    global_avg = prod_mean.mean()
    min_val    = global_avg * min_ratio if global_avg > 0 else 0

    alerts = []
    for prod in cur.index.intersection(prev.index):
        if prev[prod] <= 0 or prod_mean.get(prod, 0) < min_val:
            continue
        pct = (cur[prod] - prev[prod]) / prev[prod] * 100
        if abs(pct) >= pct_threshold:
            alerts.append({"name": prod, "prev": prev[prod], "cur": cur[prod], "pct": pct})
    return sorted(alerts, key=lambda x: abs(x["pct"]), reverse=True)


def detect_cvr_anomalies(
    cv_agg: pd.DataFrame,
    group_key: str,
    pct_threshold: float = 25.0,
    abs_threshold: float = 1.0,
    min_visits: int = 100,
) -> list[dict]:
    """직전 기간 대비 전환율 급락/급등 채널 반환
    조건: 유입수 ≥ min_visits AND 절대 변화 ≥ abs_threshold%p AND 상대 변화 ≥ pct_threshold%
    """
    dates = sorted(cv_agg["날짜"].unique())
    if len(dates) < 2:
        return []
    cur_date, prev_date = dates[-1], dates[-2]
    cur  = cv_agg[cv_agg["날짜"] == cur_date].set_index(group_key)
    prev = cv_agg[cv_agg["날짜"] == prev_date].set_index(group_key)

    alerts = []
    for ch in cur.index.intersection(prev.index):
        if cur.loc[ch, "유입수"] < min_visits:
            continue
        cvr_c, cvr_p = cur.loc[ch, "전환율"], prev.loc[ch, "전환율"]
        if cvr_p <= 0:
            continue
        abs_chg = cvr_c - cvr_p
        rel_chg = abs_chg / cvr_p * 100
        if abs(abs_chg) >= abs_threshold and abs(rel_chg) >= pct_threshold:
            alerts.append({
                "name": ch,
                "prev_cvr": cvr_p, "cur_cvr": cvr_c,
                "abs_change": abs_chg, "rel_change": rel_chg,
                "visits": cur.loc[ch, "유입수"],
            })
    return sorted(alerts, key=lambda x: abs(x["rel_change"]), reverse=True)


def _render_anomaly_cards(alerts: list[dict], kind: str = "product") -> None:
    """급락(빨강) / 급등(초록) 카드를 최대 6개, 3열로 렌더링"""
    n    = min(len(alerts), 6)
    cols = st.columns(min(3, n))
    for i, a in enumerate(alerts[:6]):
        with cols[i % 3]:
            if kind == "product":
                pct    = a["pct"]
                icon   = "🔺" if pct > 0 else "🔻"
                color  = "#f0fff4" if pct > 0 else "#fff0f0"
                border = "#28a745" if pct > 0 else "#dc3545"
                title  = f"{icon} {a['name']}"
                body   = f"{pct:+.1f}%&nbsp;&nbsp;{a['prev']:,.0f} → {a['cur']:,.0f}"
            else:  # cvr
                c      = a["abs_change"]
                icon   = "🔺" if c > 0 else "🔻"
                color  = "#f0fff4" if c > 0 else "#fff0f0"
                border = "#28a745" if c > 0 else "#dc3545"
                title  = f"{icon} {a['name']}"
                body   = (f"{c:+.2f}%p&nbsp;({a['rel_change']:+.1f}%)"
                          f"&nbsp;유입 {int(a['visits']):,}")
            st.markdown(
                f'<div style="background:{color};border-left:3px solid {border};'
                f'border-radius:6px;padding:8px 10px;margin:2px 0;font-size:0.84em;">'
                f'<b>{title}</b><br><span style="color:#444;">{body}</span></div>',
                unsafe_allow_html=True,
            )


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
    st.session_state.pop("df_traffic", None)
    st.toast("새 파일이 감지되어 대시보드를 갱신합니다!", icon="🔄")

# ── Sidebar ──────────────────────────────────────────────────────────────────
st.sidebar.title("🛒 브랜드 스토어\n판매 대시보드")
st.sidebar.markdown("---")

df_raw = pd.DataFrame()

if USE_GDRIVE:
    # ── Google Drive 모드 (Streamlit Cloud) ──────────────────────────────
    folder_id         = st.secrets["gdrive"]["folder_id"]
    folder_id_traffic = st.secrets["gdrive"].get("folder_id_traffic", "")
    creds_info        = dict(st.secrets["gdrive"]["credentials"])

    if "df_raw" not in st.session_state:
        with st.spinner("판매 데이터 로딩 중..."):
            df_raw = load_from_gdrive(folder_id, creds_info)
            st.session_state["df_raw"] = df_raw
    else:
        df_raw = st.session_state["df_raw"]

    if folder_id_traffic:
        if "df_traffic" not in st.session_state:
            with st.spinner("트래픽 데이터 로딩 중..."):
                df_traffic = load_from_gdrive(folder_id_traffic, creds_info)
                st.session_state["df_traffic"] = df_traffic
        else:
            df_traffic = st.session_state["df_traffic"]
    else:
        df_traffic = pd.DataFrame()

    if st.sidebar.button("수동 새로고침", use_container_width=True):
        st.session_state.pop("df_raw", None)
        st.session_state.pop("df_traffic", None)
        st.rerun()

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

        st.sidebar.markdown("---")
        traffic_folder = st.sidebar.text_input("트래픽 데이터 폴더 경로", value="", key="traffic_folder")
        if traffic_folder and os.path.isdir(traffic_folder):
            if "df_traffic" not in st.session_state:
                with st.spinner("트래픽 데이터 로딩 중..."):
                    df_traffic = load_from_folder(traffic_folder)
                    st.session_state["df_traffic"] = df_traffic
            else:
                df_traffic = st.session_state["df_traffic"]
        else:
            df_traffic = pd.DataFrame()
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
        df_traffic = st.session_state.get("df_traffic", pd.DataFrame())

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

_max_date = df_raw["날짜"].max().date()
_min_date = df_raw["날짜"].min().date()

# 기간 기본값 우선순위: session_state > URL 파라미터 > 최근 30일
if "date_range" in st.session_state:
    _start, _end = st.session_state["date_range"]
else:
    _qs = st.query_params.get("ds")
    _qe = st.query_params.get("de")
    if _qs and _qe:
        try:
            _start = datetime.fromisoformat(_qs).date()
            _end   = datetime.fromisoformat(_qe).date()
        except Exception:
            _start = None
    else:
        _start = None
    if _start is None:
        _start = max(_min_date, _max_date - timedelta(days=29))
        _end   = _max_date

# 실제 데이터 범위 내로 클램핑
_start = max(_min_date, min(_start, _max_date))
_end   = max(_min_date, min(_end,   _max_date))

date_range = st.sidebar.date_input("기간 선택", value=(_start, _end))
if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    st.session_state["date_range"] = tuple(date_range)
    st.query_params["ds"] = date_range[0].isoformat()
    st.query_params["de"] = date_range[1].isoformat()
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
tab_prod_trend, tab_period, tab_product, tab_traffic, tab_conversion, tab_insight = st.tabs(
    ["📈 상품 판매 추이", "📊 판매 추이", "📦 상품 분석", "📺 채널 트래픽", "🎯 전환율 분석", "💡 인사이트"]
)

with tab_period:
    period_sel = st.radio("기간 단위", ["일간", "주간", "월간"], horizontal=True, key="main_period")

    if period_sel == "일간":
        st.plotly_chart(render_trend_chart(daily, "일간 판매 추이", "%m/%d"), use_container_width=True)
        st.subheader("일별 상세 데이터")
        disp = daily.copy()
        disp["날짜"] = disp["날짜"].dt.strftime("%Y-%m-%d")
        disp["결제금액"] = disp["결제금액"].apply(lambda x: f"{int(x):,}")
        disp["환불금액"] = disp["환불금액"].apply(lambda x: f"{int(x):,}")
        disp["모바일비율"] = (disp["모바일비율"] * 100).map("{:.1f}%".format)
        st.dataframe(disp, use_container_width=True, hide_index=True)

    elif period_sel == "주간":
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

    else:
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
        pa1, pa2 = st.columns([1, 2])
        with pa1:
            pa_period = st.radio("기간 단위", ["전체", "일간", "주간", "월간"], horizontal=True, key="pa_period")

        df_pa = df_raw.copy()
        pa_title_suffix = "전체 기간"

        if pa_period != "전체":
            if pa_period == "일간":
                avail_vals   = sorted(df_pa["날짜"].dt.date.unique(), reverse=True)
                avail_labels = [str(v) for v in avail_vals]
            elif pa_period == "주간":
                df_pa["_p"]  = df_pa["날짜"].dt.to_period("W").dt.start_time
                avail_vals   = sorted(df_pa["_p"].unique(), reverse=True)
                avail_labels = [v.strftime("%Y-%m-%d 주") for v in avail_vals]
            else:
                df_pa["_p"]  = df_pa["날짜"].dt.to_period("M").dt.start_time
                avail_vals   = sorted(df_pa["_p"].unique(), reverse=True)
                avail_labels = [v.strftime("%Y-%m") for v in avail_vals]

            with pa2:
                sel_label = st.selectbox("분석 기간", avail_labels, key="pa_pick")
            sel_val = avail_vals[avail_labels.index(sel_label)]
            pa_title_suffix = sel_label

            if pa_period == "일간":
                df_pa = df_pa[df_pa["날짜"].dt.date == sel_val]
            else:
                df_pa = df_pa[df_pa["_p"] == sel_val]

        prod_agg = (
            df_pa.groupby("상품명")
            .agg(결제금액=("결제금액", "sum"), 결제수=("결제수", "sum"),
                 환불금액=("환불금액", "sum"), 결제상품수량=("결제상품수량", "sum"))
            .reset_index()
            .sort_values("결제금액", ascending=False)
        )
        prod_agg["환불율"] = (prod_agg["환불금액"] / prod_agg["결제금액"].replace(0, np.nan) * 100).fillna(0)

        top_n = st.slider("상위 N개 상품", 5, min(30, max(len(prod_agg), 5)), 10, key="pa_topn")
        top = prod_agg.head(top_n)

        col_a, col_b = st.columns(2)
        with col_a:
            fig_bar = px.bar(
                top, x="결제금액", y="상품명", orientation="h",
                title=f"상위 상품 결제금액 ({pa_title_suffix})", color="결제금액",
                color_continuous_scale="Greens",
                labels={"결제금액": "결제금액(원)", "상품명": ""},
            )
            fig_bar.update_layout(yaxis=dict(autorange="reversed"), height=400)
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_b:
            fig_scatter = px.scatter(
                prod_agg.head(top_n * 2), x="결제수", y="환불율",
                size="결제금액", hover_name="상품명",
                title=f"결제수 vs 환불율 ({pa_title_suffix})",
                labels={"결제수": "결제건수", "환불율": "환불율(%)"},
                color="환불율", color_continuous_scale="RdYlGn_r",
            )
            fig_scatter.update_layout(height=400)
            st.plotly_chart(fig_scatter, use_container_width=True)

        if "상품카테고리(대)" in df_pa.columns:
            cat_agg = (
                df_pa.groupby("상품카테고리(대)")
                .agg(결제금액=("결제금액", "sum"))
                .reset_index()
            )
            fig_pie = px.pie(cat_agg, names="상품카테고리(대)", values="결제금액",
                             title=f"카테고리별 매출 비중 ({pa_title_suffix})")
            st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader("전체 상품 상세")
        disp_prod = prod_agg.copy()
        disp_prod["결제금액"] = disp_prod["결제금액"].apply(lambda x: f"{int(x):,}")
        disp_prod["환불금액"] = disp_prod["환불금액"].apply(lambda x: f"{int(x):,}")
        disp_prod["환불율"]   = disp_prod["환불율"].map("{:.1f}%".format)
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

        # 날짜별 순위 계산 & 전체 합산 기준 정렬 (툴팁 표시 순서)
        df_grouped["순위"] = (
            df_grouped.groupby("날짜")[metric_sel]
            .rank(ascending=False, method="min")
            .astype(int)
        )
        total_rank_order = (
            df_grouped.groupby("상품명_단축")[metric_sel]
            .sum()
            .sort_values(ascending=False)
            .index.tolist()
        )

        # ── 급락/급등 감지 ────────────────────────────────────────────
        with st.expander("⚠️ 급락/급등 감지", expanded=True):
            pt_thresh = st.slider(
                "임계값 (직전 기간 대비 변화율 %)", 10, 80, 30, 5, key="pt_thresh",
                help="이 값 이상 변화한 상품만 표시합니다. 전체 평균의 10% 미만 소량 상품은 제외.",
            )
            prod_alerts = detect_product_anomalies(df_grouped, metric_sel, pct_threshold=pt_thresh)
            if prod_alerts:
                st.caption(f"직전 {period_sel} 대비 ±{pt_thresh}% 이상 변화 상품 — {len(prod_alerts)}건")
                _render_anomaly_cards(prod_alerts, kind="product")
            elif len(df_grouped["날짜"].unique()) < 2:
                st.caption("기간이 2개 이상 있어야 감지됩니다.")
            else:
                st.success(f"✅ 직전 {period_sel} 대비 ±{pt_thresh}% 이상 변화 상품 없음")

        # ── 멀티라인 차트 ──────────────────────────────────────────────
        fig_line = px.line(
            df_grouped,
            x="날짜", y=metric_sel, color="상품명_단축",
            markers=True,
            custom_data=["순위", "상품명_단축"],
            title=f"상위 {top_n_trend}개 상품 {period_sel} {metric_sel} 추이",
            labels={"상품명_단축": "상품명", metric_sel: metric_sel},
            category_orders={"상품명_단축": total_rank_order},
        )
        fig_line.update_layout(
            height=460,
            legend=dict(orientation="h", y=-0.25, font=dict(size=11)),
            xaxis_tickformat=x_fmt,
            hovermode="x unified",
            plot_bgcolor="white",
        )
        fig_line.update_traces(
            line=dict(width=2.5),
            hovertemplate="%{customdata[0]}위&nbsp;&nbsp;%{customdata[1]}: %{y:,.0f}<extra></extra>",
        )
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

with tab_traffic:
    st.subheader("채널별 트래픽 추이")
    if df_traffic.empty:
        st.info("트래픽 데이터가 없습니다. 사이드바에서 트래픽 데이터 폴더를 설정하거나 Google Drive 연동을 확인하세요.")
    else:
        TRAFFIC_COLS = {
            "유입수": "유입수", "고객수": "고객수", "광고비": "광고비", "페이지수": "페이지수",
        }
        tc1, tc2 = st.columns([1, 1])
        with tc1:
            t_metric = st.selectbox("지표", list(TRAFFIC_COLS.keys()), key="tc_metric")
        with tc2:
            t_period = st.radio("기간 단위", ["일간", "주간", "월간"], horizontal=True, key="tc_period")

        # 채널그룹 필터 (기본 표시 단위)
        all_groups = sorted(df_traffic["채널그룹"].dropna().unique()) if "채널그룹" in df_traffic.columns else []
        sel_groups = st.multiselect("채널그룹 선택 (기본 표시 단위)", all_groups, default=all_groups, key="tc_groups")

        # 채널명 세부 필터 (선택 시 채널명 단위로 드릴다운)
        df_t_base = df_traffic[df_traffic["채널그룹"].isin(sel_groups)] if sel_groups else df_traffic
        all_channels = sorted(df_t_base["채널명"].dropna().unique()) if "채널명" in df_t_base.columns else []
        sel_channels = st.multiselect(
            "채널명 세부 필터 (선택 시 채널명 단위로 표시 / 미선택 시 채널그룹 단위)",
            all_channels, default=[], key="tc_channels"
        )

        # 그룹 키 및 적용 데이터 결정
        if sel_channels:
            df_t = df_t_base[df_t_base["채널명"].isin(sel_channels)]
            t_group_key = "채널명"
        else:
            df_t = df_t_base
            t_group_key = "채널그룹"

        if t_period == "일간":
            t_grouped = df_t.groupby(["날짜", t_group_key])[t_metric].sum().reset_index()
        elif t_period == "주간":
            df_t = df_t.copy(); df_t["_기간"] = df_t["날짜"].dt.to_period("W").dt.start_time
            t_grouped = df_t.groupby(["_기간", t_group_key])[t_metric].sum().reset_index().rename(columns={"_기간": "날짜"})
        else:
            df_t = df_t.copy(); df_t["_기간"] = df_t["날짜"].dt.to_period("M").dt.start_time
            t_grouped = df_t.groupby(["_기간", t_group_key])[t_metric].sum().reset_index().rename(columns={"_기간": "날짜"})

        t_title_prefix = "채널명별" if sel_channels else "채널그룹별"
        fig_tl = px.line(t_grouped, x="날짜", y=t_metric, color=t_group_key,
                         markers=True, title=f"{t_title_prefix} {t_period} {t_metric} 추이")
        fig_tl.update_layout(height=420, hovermode="x unified",
                             legend=dict(orientation="h", y=-0.25), plot_bgcolor="white")
        st.plotly_chart(fig_tl, use_container_width=True)

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            if "채널속성" in df_t.columns:
                attr_agg = df_t.groupby("채널속성")[t_metric].sum().reset_index()
                fig_pie = px.pie(attr_agg, names="채널속성", values=t_metric,
                                 title=f"채널속성별 {t_metric} 비중", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
        with col_t2:
            top_ch = df_t.groupby("채널명")[t_metric].sum().nlargest(10).reset_index()
            fig_bar = px.bar(top_ch, x=t_metric, y="채널명", orientation="h",
                             title=f"채널명 상위 10 ({t_metric})", color=t_metric,
                             color_continuous_scale="Blues")
            fig_bar.update_layout(yaxis=dict(autorange="reversed"), height=380)
            st.plotly_chart(fig_bar, use_container_width=True)

        fig_stack = px.bar(t_grouped, x="날짜", y=t_metric, color=t_group_key,
                           barmode="stack", title=f"{t_title_prefix} {t_period} {t_metric} 구성")
        fig_stack.update_layout(height=360, legend=dict(orientation="h", y=-0.25), plot_bgcolor="white")
        st.plotly_chart(fig_stack, use_container_width=True)


with tab_conversion:
    st.subheader("채널별 전환율 & ROAS 분석")
    if df_traffic.empty:
        st.info("트래픽 데이터가 없습니다.")
    else:
        cv1, cv2 = st.columns([1, 1])
        with cv1:
            cv_period = st.radio("기간 단위", ["일간", "주간", "월간"], horizontal=False, key="cv_period")
        with cv2:
            # 채널그룹 필터 (기본 표시 단위)
            cv_all_groups = sorted(df_traffic["채널그룹"].dropna().unique()) if "채널그룹" in df_traffic.columns else []
            cv_sel_groups = st.multiselect("채널그룹 선택 (기본 표시 단위)", cv_all_groups, default=cv_all_groups, key="cv_groups")

        # 채널명 세부 필터 (선택 시 채널명 단위로 드릴다운)
        df_cv_base = df_traffic[df_traffic["채널그룹"].isin(cv_sel_groups)] if cv_sel_groups else df_traffic
        cv_all_channels = sorted(df_cv_base["채널명"].dropna().unique()) if "채널명" in df_cv_base.columns else []
        cv_sel_channels = st.multiselect(
            "채널명 세부 필터 (선택 시 채널명 단위로 표시 / 미선택 시 채널그룹 단위)",
            cv_all_channels, default=[], key="cv_channels"
        )

        if cv_sel_channels:
            df_cv = df_cv_base[df_cv_base["채널명"].isin(cv_sel_channels)]
            cv_group_key = "채널명"
        else:
            df_cv = df_cv_base
            cv_group_key = "채널그룹"

        def _period_col(d, p):
            d = d.copy()
            if p == "주간": d["_기간"] = d["날짜"].dt.to_period("W").dt.start_time
            elif p == "월간": d["_기간"] = d["날짜"].dt.to_period("M").dt.start_time
            else: d["_기간"] = d["날짜"]
            return d

        df_cv = _period_col(df_cv, cv_period)
        cv_agg = df_cv.groupby(["_기간", cv_group_key]).agg(
            유입수=("유입수", "sum"),
            결제수=("결제수(마지막클릭)", "sum"),
            결제금액=("결제금액(마지막클릭)", "sum"),
            광고비=("광고비", "sum"),
        ).reset_index().rename(columns={"_기간": "날짜"})
        cv_agg["전환율"] = (cv_agg["결제수"] / cv_agg["유입수"].replace(0, np.nan) * 100).fillna(0)
        cv_agg["ROAS"] = (cv_agg["결제금액"] / cv_agg["광고비"].replace(0, np.nan)).fillna(0)

        # KPI (필터 적용된 데이터 기준)
        total_visits = df_cv["유입수"].sum()
        total_orders_cv = df_cv["결제수(마지막클릭)"].sum()
        total_revenue_cv = df_cv["결제금액(마지막클릭)"].sum()
        total_adspend = df_cv["광고비"].sum()
        overall_cvr = total_orders_cv / total_visits * 100 if total_visits else 0
        overall_roas = total_revenue_cv / total_adspend if total_adspend else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("총 유입수", f"{int(total_visits):,}")
        k2.metric("총 결제수", f"{int(total_orders_cv):,}")
        k3.metric("전체 전환율", f"{overall_cvr:.2f}%")
        k4.metric("전체 ROAS", f"{overall_roas:.1f}" if overall_roas else "N/A")
        st.markdown("---")

        # ── 전환율 급락/급등 감지 ──────────────────────────────────────
        with st.expander("⚠️ 전환율 급락/급등 감지", expanded=True):
            cvr_c1, cvr_c2, cvr_c3 = st.columns(3)
            with cvr_c1:
                cvr_min_v = st.number_input(
                    "최소 유입수", min_value=10, max_value=10000, value=100, step=10, key="cvr_min_v",
                    help="유입수가 이 값 미만인 채널은 노이즈로 보고 제외합니다.",
                )
            with cvr_c2:
                cvr_abs_t = st.number_input(
                    "절대 변화 기준 (%p)", min_value=0.1, max_value=10.0, value=1.0, step=0.1,
                    key="cvr_abs_t", help="전환율이 이 값(%p) 이상 변해야 표시합니다.",
                )
            with cvr_c3:
                cvr_rel_t = st.slider(
                    "상대 변화 기준 (%)", 10, 80, 25, 5, key="cvr_rel_t",
                    help="두 조건(절대 + 상대)을 모두 만족해야 표시합니다.",
                )
            cvr_alerts = detect_cvr_anomalies(
                cv_agg, cv_group_key,
                pct_threshold=float(cvr_rel_t),
                abs_threshold=float(cvr_abs_t),
                min_visits=int(cvr_min_v),
            )
            if cvr_alerts:
                st.caption(
                    f"직전 {cv_period} 대비 전환율 절대 {cvr_abs_t}%p 이상 & 상대 {cvr_rel_t}% 이상 변화 — {len(cvr_alerts)}건"
                )
                _render_anomaly_cards(cvr_alerts, kind="cvr")
            elif len(cv_agg["날짜"].unique()) < 2:
                st.caption("기간이 2개 이상 있어야 감지됩니다.")
            else:
                st.success("✅ 설정 기준 이상의 전환율 변화 채널 없음")
        st.markdown("---")

        cv_title_prefix = "채널명별" if cv_sel_channels else "채널그룹별"
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            fig_cvr = px.line(cv_agg, x="날짜", y="전환율", color=cv_group_key,
                              markers=True, title=f"{cv_title_prefix} {cv_period} 전환율(%)")
            fig_cvr.update_layout(height=380, hovermode="x unified", plot_bgcolor="white",
                                  legend=dict(orientation="h", y=-0.3))
            st.plotly_chart(fig_cvr, use_container_width=True)
        with col_c2:
            cv_roas = cv_agg[cv_agg["ROAS"] > 0]
            fig_roas = px.line(cv_roas, x="날짜", y="ROAS", color=cv_group_key,
                               markers=True, title=f"{cv_title_prefix} {cv_period} ROAS")
            fig_roas.update_layout(height=380, hovermode="x unified", plot_bgcolor="white",
                                   legend=dict(orientation="h", y=-0.3))
            st.plotly_chart(fig_roas, use_container_width=True)

        # 버블차트: 항상 채널명 레벨
        ch_agg = df_cv.groupby(["채널명", "채널그룹"]).agg(
            유입수=("유입수", "sum"),
            결제수=("결제수(마지막클릭)", "sum"),
            결제금액=("결제금액(마지막클릭)", "sum"),
            광고비=("광고비", "sum"),
        ).reset_index()
        ch_agg["전환율"] = (ch_agg["결제수"] / ch_agg["유입수"].replace(0, np.nan) * 100).fillna(0)
        ch_agg["ROAS"] = (ch_agg["결제금액"] / ch_agg["광고비"].replace(0, np.nan)).fillna(0)
        ch_agg["채널명_단축"] = ch_agg["채널명"].str[:15]

        fig_bubble = px.scatter(
            ch_agg, x="유입수", y="전환율", size="결제금액",
            color="채널그룹", hover_name="채널명",
            title="채널별 유입수 vs 전환율 (버블 크기 = 결제금액)",
            labels={"유입수": "유입수", "전환율": "전환율(%)"},
        )
        fig_bubble.update_layout(height=440)
        st.plotly_chart(fig_bubble, use_container_width=True)

        ch_paid = ch_agg[ch_agg["광고비"] > 0].nlargest(15, "유입수")
        if not ch_paid.empty:
            fig_adrev = go.Figure()
            fig_adrev.add_bar(x=ch_paid["채널명_단축"], y=ch_paid["광고비"], name="광고비", marker_color="#f44336")
            fig_adrev.add_bar(x=ch_paid["채널명_단축"], y=ch_paid["결제금액"], name="결제금액", marker_color="#4CAF50")
            fig_adrev.update_layout(barmode="group", title="채널별 광고비 vs 결제금액",
                                    height=380, plot_bgcolor="white",
                                    legend=dict(orientation="h"))
            st.plotly_chart(fig_adrev, use_container_width=True)


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
