"""
온/오프라인 데일리 매출 대시보드  ·  데이터 소스: Google Spreadsheet
"""

import re
import json
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ─────────────────────────────────────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────────────────────────────────────

SPREADSHEET_ID = "1e5yROK_nj8yCeNOOyCOAth3WH3a5A6OSbqaIZjT23BE"

COL = {
    "날짜": 0, "요일": 1,
    "총_매출": 4, "총_수량": 5,
    "온라인_매출": 8,
    "오프라인_매출": 12,
    "채널_시작": 16,
    "오프라인_정규_통합_시작": 42,
    "오프라인_개별_시작": 44,
}

st.set_page_config(
    page_title="온/오프라인 매출 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .kpi-card { background:#f8f9fa; border-radius:10px; padding:16px;
              border-left:4px solid #4CAF50; margin-bottom:8px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 인증
# ─────────────────────────────────────────────────────────────────────────────

def _get_creds_info() -> dict:
    try:
        if "gdrive" in st.secrets:
            return dict(st.secrets["gdrive"]["credentials"])
    except Exception:
        pass
    return {
        "type": "service_account",
        "project_id": "brand-store-dashboard",
        "private_key_id": "ffa7200ecd737ac8b81a48523bd3c89c5f98ec39",
        "private_key": (
            "-----BEGIN PRIVATE KEY-----\n"
            "MIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCi/WyZPHjubbqY\n"
            "WT21O/bBCpgEztPwvaUlDJRH6Tf46x9HmeKuqWzi82W58S49et3QGips8+CSQVMv\n"
            "H6xiCB2nM/zXP9RrjMvQmwlalkPsiSPG0RzpUTLEjrEpXzCmPdTL4mWUDgLGj2QQ\n"
            "Zi8yBvq1K0wBD1PwBoFOlQ7epHZe/nuN8Rp6Shk9lXCp5rxQy7G+7+PsmO8ofq1O\n"
            "4VQsFWN1haQdmV2e/3na0GSja4yjfWdK+MkHROizJaCzOkbCqcoD+PeW5MZCQCsE\n"
            "SFmMdn8dHPsuGKEsuySOfs0ctUfcPSchZ8BRkJ3+IMXcylSXn//ZRFXt5NXzt104\n"
            "ZbxsdmXHAgMBAAECggEAFw2cLiJZAnQ/t6urqDYmhRBQByIP5RWVxaM9V9x7P7hV\n"
            "0MVJxCWPkpwVj9K2lRiLRhNlW8q0wUr/DoO+2JDITO1z/if3jgy4iTVcgdUOH2cY\n"
            "Sdcq7S9AbH+4PB9MHL9TGVFYx7Ohnd8LKqyhUUSeckCA/AdkemcZI/m+RCxj17SS\n"
            "EQtrkM9xUksIS8sjVtwFGrXKBZSTGyckqQLAUMA+zEVnebT00Mee7jk2JD3BcdSG\n"
            "SXnZx6c8dQgBt+6ajo0axkdAcQN5jgArrRCBSkDn49jVp3/DiD8uOzC2NQsPFtkL\n"
            "WkT9jH+xWqRvNF7slpEX1/Kmq1jRIIVFnOfH4AWhAQKBgQDN8IFrpUNyjnJ9AcGi\n"
            "c5ud5ga0NoCBtDWQ+MDwrqzKVtD37WXXgbiZA/2IkmSYRTu6AzwsPdHUpY4EI96W\n"
            "Y1BIBi1e8t7J9ArG5kLPSj4G+8FCM6y65Ua1Echivg+ZukI00WdEuyF/z2Pf9OnT\n"
            "udMgPX07rBo7rm3cABo7eohBAQKBgQDKnDGfm8ngERJyqRXy/hWGO4IVgYXe0U7f\n"
            "xGC6cYD4HBcfPVJKAyOCIR/f6BnW+xjQSQt3jvgYQeMKr7B2BrwNOl+bHnQw7mW7\n"
            "Bd18Jx9Y41LCu1ToOdHaJjrmNpEhELq+hWeR5yAmXKTY9WRkRM7yf2ouLtwrBKpC\n"
            "QtHzly3exwKBgDn5F2XHOyp3gTFBmlHx+3/CrmZy5VAd++pYrG/UrF21fNQeZ0n4\n"
            "gY/JuMiGdX0MGFkv6fOGX5heFpGBy3pIcOQloQYWlrMBWTtOvMX/32A15NyPEXP3\n"
            "cSUt4Vwyps+eyF54CHsntrF1H2d/WYe5yv5LcQKoWyYr309MVBYkU1EBAoGAF6jX\n"
            "rqC9mTnFIriWBJMhJlSqoyJF5LgicsT22q7IdbCqDo7Vnijxq498rmPnKJCX3DK7\n"
            "cRGz7Pk8rxHHFHFC4nSPl4id3tzn6kgMDiRvZ6zcDDtd9eRSmhvewuVaWzcd54Oi\n"
            "jYok2fX1lhRJzd+vHug8GPqF4UwhKa2t4LkltR8CgYAfZOsd8lj49A87G6nbVItV\n"
            "af3vxksasZY907aDVrYlIYHlliUNyOKxJMbtzM9rz++YMadJsMyDBQJz1h/Xd4rA\n"
            "lj3qiCBoSBrMTCSBnVSxV2PYfkjAvphACeezuzDJxly9zYdCTRz/bVWyouCNMJ/l\n"
            "paM05fOfW9YcqgKv23rCVg==\n"
            "-----END PRIVATE KEY-----\n"
        ),
        "client_email": "dashboard-reader@brand-store-dashboard.iam.gserviceaccount.com",
        "client_id": "116554768869865545577",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": (
            "https://www.googleapis.com/robot/v1/metadata/x509/"
            "dashboard-reader%40brand-store-dashboard.iam.gserviceaccount.com"
        ),
    }


# ── Sheets 서비스 객체: 세션 간 재사용 (build 1회만 호출) ─────────────────
@st.cache_resource
def _get_sheets_service(creds_json: str):
    """Google Sheets API 서비스 — 세션/재실행 간 재사용"""
    creds = service_account.Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _fetch(svc, sheet: str, range_: str) -> list:
    res = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet}'!{range_}",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()
    return res.get("values", [])


# ─────────────────────────────────────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────────────────────────────────────

def _num(v) -> float:
    s = str(v).replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _is_daily(row) -> bool:
    if not row:
        return False
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}$", str(row[0]).strip()))


def _pad(row: list, length: int) -> list:
    return row + [""] * max(0, length - len(row))


def fmt_won(v: float) -> str:
    if v >= 1_000_000_000:
        return f"{v/1_000_000_000:.1f}B원"
    if v >= 1_000_000:
        return f"{v/1_000_000:.1f}M원"
    if v >= 1_000:
        return f"{v/1_000:.0f}K원"
    return f"{int(v):,}원"


# ─────────────────────────────────────────────────────────────────────────────
# 채널 메타 파싱
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(ttl=600, show_spinner=False)
def load_channel_meta(creds_json: str) -> list[dict]:
    svc = _get_sheets_service(creds_json)
    rows = _fetch(svc, "통합_채널별", "A1:CZ4")
    max_col = max((len(r) for r in rows if r), default=0)

    def pad(r):
        return _pad(r, max_col) if r else [""] * max_col

    row2 = pad(rows[1] if len(rows) > 1 else [])
    row4 = pad(rows[3] if len(rows) > 3 else [])

    online_start      = COL["채널_시작"]
    offline_agg_start = COL["오프라인_정규_통합_시작"]
    offline_indiv_start = COL["오프라인_개별_시작"]

    channels: list[dict] = []
    cur: dict | None = None

    for i in range(online_start, max_col):
        ch_raw = row2[i].replace("\n", " ").strip()
        metric = row4[i].replace("\n", " ").strip()

        if ch_raw and ch_raw not in ("채널", "수수료"):
            if cur:
                channels.append(cur)
            if i < offline_agg_start:
                ch_type = "온라인"
            elif i < offline_indiv_start:
                ch_type = "오프라인_정규통합"
            else:
                ch_type = "오프라인"
            cur = {"name": ch_raw, "type": ch_type, "sales_idx": i, "qty_idx": -1}
        elif cur and cur["qty_idx"] == -1 and "수량" in metric:
            cur["qty_idx"] = i

    if cur:
        channels.append(cur)
    return channels


# ─────────────────────────────────────────────────────────────────────────────
# 데이터 로딩: 통합_채널별
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(ttl=300, show_spinner="채널별 데이터 로딩 중…")
def load_channel_df(creds_json: str) -> pd.DataFrame:
    svc = _get_sheets_service(creds_json)
    channels = load_channel_meta(creds_json)
    max_col = max((ch["sales_idx"] for ch in channels), default=78) + 2

    raw = _fetch(svc, "통합_채널별", "A5:CZ600")
    today = datetime.now().date()
    records = []
    for row in raw:
        if not _is_daily(row):
            continue
        try:
            row_date = datetime.strptime(row[0].strip(), "%Y-%m-%d").date()
            if row_date > today:
                continue
        except Exception:
            continue
        row = _pad(row, max_col)
        # 총_매출, 온라인_매출, 오프라인_매출 모두 0이면 미입력 행
        if (_num(row[COL["총_매출"]]) == 0
                and _num(row[COL["온라인_매출"]]) == 0
                and _num(row[COL["오프라인_매출"]]) == 0):
            continue
        rec: dict = {
            "날짜":          pd.to_datetime(row[COL["날짜"]].strip()),
            "요일":          row[COL["요일"]].strip(),
            "총_매출":       _num(row[COL["총_매출"]]),
            "총_수량":       _num(row[COL["총_수량"]]),
            "온라인_매출":   _num(row[COL["온라인_매출"]]),
            "오프라인_매출": _num(row[COL["오프라인_매출"]]),
        }
        for ch in channels:
            si, qi, nm = ch["sales_idx"], ch["qty_idx"], ch["name"]
            rec[f"{nm}_매출"] = _num(row[si]) if si < len(row) else 0.0
            if qi != -1:
                rec[f"{nm}_수량"] = _num(row[qi]) if qi < len(row) else 0.0
        records.append(rec)

    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values("날짜").reset_index(drop=True)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 데이터 로딩: 제품별 (3개 시트를 1번 API 세트로 통합 캐시)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_product_sheet(svc, sheet: str, today) -> pd.DataFrame:
    """제품별 시트 1개 파싱 (캐시 없음 — load_all_products 내부에서 호출)"""
    header_rows = _fetch(svc, sheet, "A4:BZ4")
    if not header_rows:
        return pd.DataFrame()
    headers = header_rows[0]
    raw = _fetch(svc, sheet, "A5:BZ600")
    records = []
    for row in raw:
        if not _is_daily(row):
            continue
        try:
            row_date = datetime.strptime(row[0].strip(), "%Y-%m-%d").date()
            if row_date > today:
                continue
        except Exception:
            continue
        row = _pad(row, len(headers))
        rec: dict = {"날짜": pd.to_datetime(row[0].strip())}
        for i, h in enumerate(headers):
            if i < 2 or not h.strip():
                continue
            rec[h.replace("\n", " ").strip()] = _num(row[i])
        records.append(rec)
    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values("날짜").reset_index(drop=True)
    return df


@st.cache_data(ttl=300, show_spinner="제품별 데이터 로딩 중…")
def load_all_products(creds_json: str) -> dict[str, pd.DataFrame]:
    """통합/온라인/오프라인 제품별 시트를 1회 인증으로 순차 로드"""
    svc = _get_sheets_service(creds_json)
    today = datetime.now().date()
    return {
        "통합":    _parse_product_sheet(svc, "통합_제품별",    today),
        "온라인":  _parse_product_sheet(svc, "온라인_제품별",  today),
        "오프라인": _parse_product_sheet(svc, "오프라인_제품별", today),
    }


# ─────────────────────────────────────────────────────────────────────────────
# 데이터 로딩: 판매 순위 (2개 시트를 1번 인증으로)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_ranking_sheet(svc, sheet: str) -> pd.DataFrame:
    h2  = _fetch(svc, sheet, "A2:CZ2")
    h3  = _fetch(svc, sheet, "A3:CZ3")
    h4  = _fetch(svc, sheet, "A4:CZ4")
    raw = _fetch(svc, sheet, "A5:CZ300")
    if not h4:
        return pd.DataFrame()

    row2 = _pad(h2[0] if h2 else [], 100)
    row3 = _pad(h3[0] if h3 else [], 100)

    records = []
    cur_date, cur_period = "", ""
    for row in raw:
        if not row:
            continue
        row = _pad(row, 100)
        if row[0].strip():
            cur_date = row[0].strip()
        if row[1].strip():
            cur_period = row[1].strip()
        rank = row[2].strip()
        if not rank.isdigit():
            continue
        i = 3
        while i < len(row2) - 1:
            ch_raw = row2[i].strip()
            sub    = row3[i].strip() if i < len(row3) else ""
            label  = f"{ch_raw}_{sub}" if sub else ch_raw
            prod   = row[i].strip()   if i < len(row) else ""
            qty    = row[i+1].strip() if i+1 < len(row) else ""
            if prod and prod not in ("세트 판매 없음",):
                records.append({
                    "기준일": cur_date,
                    "기간":   cur_period,
                    "순위":   int(rank),
                    "채널":   label,
                    "제품명": prod,
                    "판매량": _num(qty),
                })
            i += 2
    return pd.DataFrame(records) if records else pd.DataFrame()


@st.cache_data(ttl=300, show_spinner=False)
def load_all_rankings(creds_json: str) -> dict[str, pd.DataFrame]:
    """월간/주간 판매 순위를 1회 인증으로 로드"""
    svc = _get_sheets_service(creds_json)
    return {
        "월간": _parse_ranking_sheet(svc, "월간 판매 순위"),
        "주간": _parse_ranking_sheet(svc, "주간 판매 순위"),
    }


# ─────────────────────────────────────────────────────────────────────────────
# 집계 헬퍼
# ─────────────────────────────────────────────────────────────────────────────

def resample_df(df: pd.DataFrame, unit: str, agg_cols: list[str]) -> pd.DataFrame:
    if unit == "일간":
        return df
    d = df.copy()
    if unit == "주간":
        d["_기간"] = d["날짜"].dt.to_period("W").dt.start_time
    else:
        d["_기간"] = d["날짜"].dt.to_period("M").dt.start_time
    return (
        d.groupby("_기간")[agg_cols].sum()
        .reset_index()
        .rename(columns={"_기간": "날짜"})
    )


def channel_long(df: pd.DataFrame, channels: list[dict],
                 ch_type: str | None = None, unit: str = "일간",
                 metric: str = "매출") -> pd.DataFrame:
    """채널별 wide → long (pandas melt — iterrows 없음)"""
    chs = [ch for ch in channels if ch_type is None or ch["type"] == ch_type]
    col_names = [f"{ch['name']}_{metric}" for ch in chs
                 if f"{ch['name']}_{metric}" in df.columns]
    if not col_names:
        return pd.DataFrame()

    d = df[["날짜"] + col_names].copy()
    if unit == "주간":
        d["날짜"] = d["날짜"].dt.to_period("W").dt.start_time
    elif unit == "월간":
        d["날짜"] = d["날짜"].dt.to_period("M").dt.start_time
    d = d.groupby("날짜")[col_names].sum().reset_index()

    long = d.melt(id_vars="날짜", var_name="채널", value_name=metric)
    long["채널"] = long["채널"].str.replace(f"_{metric}$", "", regex=True)
    return long[long[metric] > 0]


# ─────────────────────────────────────────────────────────────────────────────
# 사이드바 & 인증
# ─────────────────────────────────────────────────────────────────────────────

st.sidebar.title("📊 온/오프라인 매출\n대시보드")
st.sidebar.markdown("---")

creds_info = _get_creds_info()
creds_json = json.dumps(creds_info)

# ── 초기 로딩: channels + channel_df 먼저, 나머지는 캐시 통합 ─────────────
with st.spinner("데이터 연결 중…"):
    try:
        channels     = load_channel_meta(creds_json)
        df_ch        = load_channel_df(creds_json)
        prod_dict    = load_all_products(creds_json)
        ranking_dict = load_all_rankings(creds_json)
    except Exception as e:
        st.error(f"데이터 로드 실패: {e}")
        st.stop()

df_prod     = prod_dict["통합"]
df_prod_on  = prod_dict["온라인"]
df_prod_off = prod_dict["오프라인"]
df_rank_m   = ranking_dict["월간"]
df_rank_w   = ranking_dict["주간"]

if df_ch.empty:
    st.warning("불러온 일별 데이터가 없습니다.")
    st.stop()

# ── 기간 필터 ────────────────────────────────────────────────────────────────
_min_d = df_ch["날짜"].min().date()
_max_d = df_ch["날짜"].max().date()

st.sidebar.subheader("기간 필터")
st.sidebar.caption(f"데이터 범위: {_min_d} ~ **{_max_d}**")

_def_start = max(_min_d, _max_d - timedelta(days=29))
date_range = st.sidebar.date_input(
    "기간 선택", value=(_def_start, _max_d),
    min_value=_min_d, max_value=_max_d,
)
if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    s_date, e_date = date_range
    df   = df_ch[(df_ch["날짜"].dt.date >= s_date) & (df_ch["날짜"].dt.date <= e_date)].copy()
    df_p  = df_prod[(df_prod["날짜"].dt.date >= s_date) & (df_prod["날짜"].dt.date <= e_date)].copy()     if not df_prod.empty     else df_prod
    df_po = df_prod_on[(df_prod_on["날짜"].dt.date >= s_date) & (df_prod_on["날짜"].dt.date <= e_date)].copy() if not df_prod_on.empty  else df_prod_on
    df_pf = df_prod_off[(df_prod_off["날짜"].dt.date >= s_date) & (df_prod_off["날짜"].dt.date <= e_date)].copy() if not df_prod_off.empty else df_prod_off
else:
    df, df_p, df_po, df_pf = df_ch, df_prod, df_prod_on, df_prod_off

if df.empty:
    st.warning("선택 기간에 데이터가 없습니다.")
    st.stop()

if st.sidebar.button("🔄 새로고침", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption("💡 데이터는 5분마다 자동 갱신됩니다")

online_channels  = [ch for ch in channels if ch["type"] == "온라인"]
offline_channels = [ch for ch in channels if ch["type"] == "오프라인"]

# ─────────────────────────────────────────────────────────────────────────────
# 메인 타이틀 & KPI (달성률 제거 — 매출 3개만)
# ─────────────────────────────────────────────────────────────────────────────

st.title("📊 온/오프라인 매출 대시보드")

total_rev   = df["총_매출"].sum()
online_rev  = df["온라인_매출"].sum()
offline_rev = df["오프라인_매출"].sum()

def _vs_prev(df_: pd.DataFrame, col: str) -> float | None:
    if len(df_) < 2:
        return None
    last, prev = df_.iloc[-1][col], df_.iloc[-2][col]
    return (last - prev) / prev * 100 if prev else None

rev_delta = _vs_prev(df, "총_매출")
on_delta  = _vs_prev(df, "온라인_매출")
off_delta = _vs_prev(df, "오프라인_매출")

k1, k2, k3 = st.columns(3)
k1.metric("📦 총 매출",      fmt_won(total_rev),
          f"{rev_delta:+.1f}%" if rev_delta is not None else None)
k2.metric("🔵 온라인 매출",  fmt_won(online_rev),
          f"{on_delta:+.1f}%"  if on_delta  is not None else None)
k3.metric("🟠 오프라인 매출", fmt_won(offline_rev),
          f"{off_delta:+.1f}%" if off_delta is not None else None)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# 탭
# ─────────────────────────────────────────────────────────────────────────────

tab_overview, tab_channel, tab_product, tab_ranking = st.tabs([
    "📈 매출 추이", "📡 채널별 분석", "📦 제품별 성과", "🏆 판매 순위",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 : 매출 추이
# ══════════════════════════════════════════════════════════════════════════════

with tab_overview:
    unit = st.radio("기간 단위", ["일간", "주간", "월간"], horizontal=True, key="ov_unit")

    agg_cols = ["총_매출", "온라인_매출", "오프라인_매출", "총_수량"]
    agg = resample_df(df[["날짜"] + agg_cols], unit, agg_cols)

    # 온/오프 매출 추이 단일 차트
    x = agg["날짜"]
    fig_trend = go.Figure()
    fig_trend.add_trace(go.Bar(x=x, y=agg["총_매출"], name="총합",
                               marker_color="#E0E0E0", opacity=0.4))
    fig_trend.add_trace(go.Scatter(x=x, y=agg["온라인_매출"], name="온라인",
                                   mode="lines+markers",
                                   line=dict(color="#2196F3", width=2.5)))
    fig_trend.add_trace(go.Scatter(x=x, y=agg["오프라인_매출"], name="오프라인",
                                   mode="lines+markers",
                                   line=dict(color="#FF9800", width=2.5)))
    fig_trend.update_layout(
        title="온라인 vs 오프라인 매출 추이",
        height=400,
        legend=dict(orientation="h", y=-0.15),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
        yaxis=dict(tickformat=","),
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    # 채널 구성 파이
    col_pie1, col_pie2 = st.columns(2)

    on_rev_by_ch = {
        ch["name"]: df[f"{ch['name']}_매출"].sum()
        for ch in online_channels if f"{ch['name']}_매출" in df.columns
    }
    with col_pie1:
        if on_rev_by_ch:
            pie_on = pd.DataFrame(on_rev_by_ch.items(), columns=["채널", "매출"])
            pie_on = pie_on[pie_on["매출"] > 0].sort_values("매출", ascending=False)
            fig_pie_on = px.pie(pie_on, names="채널", values="매출",
                                title="온라인 채널 매출 구성", hole=0.4)
            fig_pie_on.update_traces(textposition="inside")
            fig_pie_on.update_layout(height=380, legend=dict(orientation="v", x=1.0))
            st.plotly_chart(fig_pie_on, use_container_width=True)

    off_rev_by_ch = {
        ch["name"]: df[f"{ch['name']}_매출"].sum()
        for ch in offline_channels if f"{ch['name']}_매출" in df.columns
    }
    with col_pie2:
        if off_rev_by_ch:
            pie_off = pd.DataFrame(off_rev_by_ch.items(), columns=["채널", "매출"])
            pie_off = pie_off[pie_off["매출"] > 0].nlargest(12, "매출")
            fig_pie_off = px.pie(pie_off, names="채널", values="매출",
                                 title="오프라인 매장 매출 구성 (Top 12)", hole=0.4)
            fig_pie_off.update_traces(textposition="inside")
            fig_pie_off.update_layout(height=380, legend=dict(orientation="v", x=1.0))
            st.plotly_chart(fig_pie_off, use_container_width=True)

    # 일별 상세 테이블
    st.subheader("일별 상세 데이터")
    disp = agg[["날짜", "온라인_매출", "오프라인_매출", "총_매출"]].copy()
    for c in ["온라인_매출", "오프라인_매출", "총_매출"]:
        disp[c] = disp[c].apply(lambda v: f"{int(v):,}")
    disp["날짜"] = disp["날짜"].apply(
        lambda d: d.strftime("%Y-%m-%d") if unit == "일간"
        else d.strftime("%Y-%m-%d 주") if unit == "주간"
        else d.strftime("%Y-%m"))
    st.dataframe(disp, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 : 채널별 분석
# ══════════════════════════════════════════════════════════════════════════════

with tab_channel:
    ch_unit  = st.radio("기간 단위", ["일간", "주간", "월간"], horizontal=True, key="ch_unit")
    ch_focus = st.radio("채널 구분", ["온라인", "오프라인", "전체"], horizontal=True, key="ch_focus")

    # 온라인
    if ch_focus in ("온라인", "전체"):
        st.subheader("🔵 온라인 채널별 매출")
        on_long = channel_long(df, online_channels, "온라인", ch_unit, "매출")
        if not on_long.empty:
            fig_on_line = px.line(
                on_long, x="날짜", y="매출", color="채널",
                markers=True, title=f"온라인 채널별 {ch_unit} 매출",
            )
            fig_on_line.update_layout(height=380, hovermode="x unified",
                                      legend=dict(orientation="h", y=-0.25),
                                      plot_bgcolor="white")
            st.plotly_chart(fig_on_line, use_container_width=True)

            fig_on_bar = px.bar(
                on_long, x="날짜", y="매출", color="채널", barmode="stack",
                title=f"온라인 채널 구성 ({ch_unit})",
            )
            fig_on_bar.update_layout(height=320,
                                     legend=dict(orientation="h", y=-0.3),
                                     plot_bgcolor="white")
            st.plotly_chart(fig_on_bar, use_container_width=True)

            on_total = on_long.groupby("채널")["매출"].sum().reset_index().sort_values("매출", ascending=True)
            fig_on_tot = px.bar(on_total, x="매출", y="채널", orientation="h",
                                title="온라인 채널별 총 매출 (기간 합산)",
                                color="매출", color_continuous_scale="Blues")
            fig_on_tot.update_layout(height=max(300, len(on_total)*32+80),
                                     yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_on_tot, use_container_width=True)

    # 오프라인
    if ch_focus in ("오프라인", "전체"):
        st.subheader("🟠 오프라인 매장별 매출")
        off_long = channel_long(df, offline_channels, "오프라인", ch_unit, "매출")
        if not off_long.empty:
            top_n = st.slider("상위 N 매장", 5, min(20, len(offline_channels)), 10, key="off_topn")
            top_stores = (off_long.groupby("채널")["매출"].sum()
                          .nlargest(top_n).index.tolist())
            off_top = off_long[off_long["채널"].isin(top_stores)]

            fig_off_bar = px.bar(
                off_top, x="날짜", y="매출", color="채널", barmode="stack",
                title=f"오프라인 상위 {top_n} 매장 {ch_unit} 매출",
            )
            fig_off_bar.update_layout(height=380,
                                      legend=dict(orientation="h", y=-0.3),
                                      plot_bgcolor="white")
            st.plotly_chart(fig_off_bar, use_container_width=True)

            off_total = (off_long.groupby("채널")["매출"].sum()
                         .reset_index().nlargest(top_n, "매출")
                         .sort_values("매출", ascending=True))
            fig_off_tot = px.bar(off_total, x="매출", y="채널", orientation="h",
                                 title=f"오프라인 매장별 총 매출 Top{top_n}",
                                 color="매출", color_continuous_scale="Oranges")
            fig_off_tot.update_layout(height=max(300, top_n*32+80),
                                      yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_off_tot, use_container_width=True)

    # 기간 비교
    st.markdown("---")
    st.subheader("📊 기간 비교")
    _tmax = df_ch["날짜"].max().date()
    _tmin = df_ch["날짜"].min().date()
    comp_c1, comp_c2 = st.columns(2)
    with comp_c1:
        cur_range = st.date_input("현재 기간", (_tmax - timedelta(days=13), _tmax),
                                  min_value=_tmin, max_value=_tmax, key="comp_cur")
    with comp_c2:
        prev_range = st.date_input("비교 기간",
                                   (_tmax - timedelta(days=27), _tmax - timedelta(days=14)),
                                   min_value=_tmin, max_value=_tmax, key="comp_prev")

    if isinstance(cur_range, tuple) and isinstance(prev_range, tuple):
        df_cur  = df_ch[(df_ch["날짜"].dt.date >= cur_range[0])  & (df_ch["날짜"].dt.date <= cur_range[1])]
        df_prev = df_ch[(df_ch["날짜"].dt.date >= prev_range[0]) & (df_ch["날짜"].dt.date <= prev_range[1])]
        target_chs = online_channels if ch_focus == "온라인" else \
                     offline_channels if ch_focus == "오프라인" else \
                     online_channels + offline_channels
        comp_rows = []
        for ch in target_chs:
            col = f"{ch['name']}_매출"
            if col not in df_ch.columns:
                continue
            c, p = df_cur[col].sum(), df_prev[col].sum()
            comp_rows.append({"채널": ch["name"], "구분": ch["type"],
                               "현재": c, "비교": p,
                               "변화율": (c-p)/p*100 if p else 0})
        if comp_rows:
            comp_df = pd.DataFrame(comp_rows).sort_values("현재", ascending=False)
            m = comp_df.melt(id_vars="채널", value_vars=["현재", "비교"],
                             var_name="기간", value_name="매출")
            fig_comp = px.bar(m, x="채널", y="매출", color="기간", barmode="group",
                              title="채널별 기간 비교",
                              color_discrete_map={"현재": "#1565C0", "비교": "#90CAF9"})
            fig_comp.update_layout(height=420, plot_bgcolor="white",
                                   legend=dict(orientation="h", y=-0.12))
            st.plotly_chart(fig_comp, use_container_width=True)
            comp_disp = comp_df.copy()
            comp_disp["현재"]   = comp_disp["현재"].apply(lambda v: f"{int(v):,}")
            comp_disp["비교"]   = comp_disp["비교"].apply(lambda v: f"{int(v):,}")
            comp_disp["변화율"] = comp_disp["변화율"].map("{:+.1f}%".format)
            st.dataframe(comp_disp, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 : 제품별 성과
# ══════════════════════════════════════════════════════════════════════════════

with tab_product:
    prod_focus = st.radio("데이터 기준", ["온/오프 통합", "온라인만", "오프라인만"],
                          horizontal=True, key="prod_focus")
    prod_unit  = st.radio("기간 단위", ["일간", "주간", "월간"],
                          horizontal=True, key="prod_unit")

    src_df = {"온/오프 통합": df_p, "온라인만": df_po, "오프라인만": df_pf}[prod_focus]

    if src_df.empty:
        st.info("제품 데이터가 없습니다.")
    else:
        META_COLS = {"날짜", "요일", "목표", "달성률", "매출(거래액)",
                     "매출(결제액)\n*취소 제외", "매출(결제액) *취소 제외",
                     "매출 거래액+결제액", "수량",
                     "목표 *결제액 기준", "목표\n*결제액 기준",
                     "매출\n거래액+결제액", "목표\n"}
        prod_cols = [c for c in src_df.columns
                     if c not in META_COLS
                     and not c.startswith("매출")
                     and c != "날짜"]

        prod_total = src_df[prod_cols].sum().sort_values(ascending=False)
        prod_total = prod_total[prod_total > 0]

        if prod_total.empty:
            st.info("판매 데이터가 없습니다.")
        else:
            top_n_prod = st.slider("상위 N 제품", 5, min(30, len(prod_total)), 15, key="prod_n")
            top_prods  = prod_total.nlargest(top_n_prod).index.tolist()

            col_a, col_b = st.columns(2)
            with col_a:
                fig_bar = px.bar(
                    x=prod_total[top_prods].values,
                    y=[p[:20] for p in top_prods],
                    orientation="h",
                    title=f"상위 {top_n_prod} 제품 판매량 (기간 합산)",
                    color=prod_total[top_prods].values,
                    color_continuous_scale="Greens",
                    labels={"x": "판매량", "y": ""},
                )
                fig_bar.update_layout(yaxis=dict(autorange="reversed"), height=480)
                st.plotly_chart(fig_bar, use_container_width=True, key="prod_bar")

            with col_b:
                fig_pie = px.pie(
                    names=[p[:20] for p in top_prods],
                    values=prod_total[top_prods].values,
                    title=f"판매량 구성 (Top {top_n_prod})",
                    hole=0.4,
                )
                fig_pie.update_layout(height=480)
                st.plotly_chart(fig_pie, use_container_width=True, key="prod_pie")

            # 추이 차트
            st.subheader("제품별 판매 추이")
            d_trend = src_df[["날짜"] + top_prods].copy()
            if prod_unit == "주간":
                d_trend["날짜"] = d_trend["날짜"].dt.to_period("W").dt.start_time
            elif prod_unit == "월간":
                d_trend["날짜"] = d_trend["날짜"].dt.to_period("M").dt.start_time
            d_trend = d_trend.groupby("날짜")[top_prods].sum().reset_index()
            d_long  = d_trend.melt(id_vars="날짜", var_name="제품", value_name="판매량")
            d_long["제품_단축"] = d_long["제품"].str[:18]

            fig_line = px.line(d_long, x="날짜", y="판매량", color="제품_단축",
                               markers=True, title=f"상위 {top_n_prod} 제품 {prod_unit} 판매 추이")
            fig_line.update_layout(height=420, hovermode="x unified",
                                   legend=dict(orientation="h", y=-0.3, font=dict(size=10)),
                                   plot_bgcolor="white")
            st.plotly_chart(fig_line, use_container_width=True, key="prod_line")

            # 히트맵
            pivot = d_trend.set_index("날짜")[top_prods].T
            pivot.index = [p[:20] for p in pivot.index]
            pivot.columns = [c.strftime("%m/%d") if prod_unit == "일간"
                             else c.strftime("%m/%d 주") if prod_unit == "주간"
                             else c.strftime("%Y-%m")
                             for c in pivot.columns]
            fig_heat = go.Figure(go.Heatmap(
                z=pivot.values, x=list(pivot.columns), y=list(pivot.index),
                colorscale="Greens",
                hovertemplate="날짜: %{x}<br>제품: %{y}<br>판매량: %{z:,.0f}<extra></extra>",
            ))
            fig_heat.update_layout(title=f"제품 × 날짜 히트맵 ({prod_unit} 판매량)",
                                   height=max(300, top_n_prod*28+100),
                                   margin=dict(l=200))
            st.plotly_chart(fig_heat, use_container_width=True, key="prod_heat")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 : 판매 순위
# ══════════════════════════════════════════════════════════════════════════════

with tab_ranking:
    rank_unit = st.radio("순위 기준", ["월간", "주간"], horizontal=True, key="rank_unit")
    df_rank = df_rank_m if rank_unit == "월간" else df_rank_w

    if df_rank.empty:
        st.info("판매 순위 데이터가 없습니다.")
    else:
        periods = sorted(df_rank["기준일"].unique(), reverse=True)
        period_labels = df_rank.drop_duplicates("기준일").set_index("기준일")["기간"].to_dict()
        active_periods = (
            df_rank.groupby("기준일")["판매량"].sum()
            .loc[lambda s: s > 0]
            .index.tolist()
        )
        default_period = sorted(active_periods, reverse=True)[0] if active_periods else periods[0]
        default_idx = periods.index(default_period) if default_period in periods else 0
        sel_period = st.selectbox(
            "기간 선택", options=periods, index=default_idx,
            format_func=lambda x: period_labels.get(x, x),
        )

        df_sel = df_rank[df_rank["기준일"] == sel_period]
        avail_channels = sorted(df_sel["채널"].unique())
        sel_channels = st.multiselect("채널 선택", avail_channels,
                                      default=["온&오프라인 통합_단품", "온라인_단품", "오프라인_단품"],
                                      key="rank_channels")
        if not sel_channels:
            sel_channels = avail_channels[:3]

        df_view = df_sel[df_sel["채널"].isin(sel_channels)]
        n_cols  = min(3, len(sel_channels))
        cols_r  = st.columns(n_cols)
        for i, ch in enumerate(sel_channels):
            df_ch_rank = df_view[df_view["채널"] == ch].sort_values("순위")
            with cols_r[i % n_cols]:
                st.markdown(f"**{ch}**")
                if df_ch_rank.empty:
                    st.caption("데이터 없음")
                    continue
                fig_r = px.bar(
                    df_ch_rank.head(5),
                    x="판매량", y="제품명", orientation="h",
                    color="판매량", color_continuous_scale="Blues",
                    text="판매량",
                )
                fig_r.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                fig_r.update_layout(
                    height=300, showlegend=False,
                    yaxis=dict(autorange="reversed",
                               ticktext=[p[:15] for p in df_ch_rank.head(5)["제품명"]],
                               tickvals=df_ch_rank.head(5)["제품명"].tolist()),
                    xaxis_title="", yaxis_title="",
                    margin=dict(l=120, r=20, t=10, b=10),
                    plot_bgcolor="white",
                )
                st.plotly_chart(fig_r, use_container_width=True, key=f"rank_chart_{i}")
                disp_r = df_ch_rank.head(5)[["순위", "제품명", "판매량"]].copy()
                disp_r["판매량"] = disp_r["판매량"].apply(lambda v: f"{int(v):,}")
                st.dataframe(disp_r, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("전 채널 Top 5 종합")
        pivot_rank = df_view.pivot_table(
            index=["순위", "제품명"], columns="채널", values="판매량",
            aggfunc="sum",
        ).reset_index()
        st.dataframe(pivot_rank, use_container_width=True, hide_index=True)
