# ==============================================================================
# [유통 데이터 분석 시스템 v1.6] - Dashboard Application
# Copyright (c) 2026 [KACKAC]. All rights reserved.
# 본 프로그램의 핵심 비즈니스 로직(특수 출하 데이터 분류 및 수요 예측 알고리즘)은
# 저작권법의 보호를 받는 고유 자산입니다. 무단 복제 및 배포를 금합니다.
# [Open Source Notice]
# 본 소프트웨어는 아래의 오픈소스 라이브러리를 사용하여 개발되었습니다:
# - Streamlit (Apache License 2.0)
# - DuckDB (MIT License)
# - Pandas (BSD 3-Clause License)
# - Plotly (MIT License)
# ==============================================================================

"""
app.py - 유통 데이터 분석 대시보드
라이선스: streamlit(Apache 2.0), duckdb(MIT), pandas(BSD), plotly(MIT), xlsxwriter(BSD)
모두 상업적 무료 사용 가능
실행: streamlit run app.py
"""

import io
import os
import logging
import duckdb
import pandas as pd
import plotly.express as px
import streamlit as st
from datetime import datetime, timedelta

# ── 로깅 ─────────────────────────────────────────────────────────────────────
logging.basicConfig(
    filename="app_log.txt",
    level=logging.ERROR,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)

# ── 경로 및 상수 ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="유통 데이터 분석 시스템 v1.6", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, 'distribution_data.db')

ALL_MOVE_DESCS = [
    '매입', '일반 출하', '특수 출하', '출하 취소', '이월 매출',
    '지점 이동', '일반 반품', 'RTP', 'POST', '스크랩', '재고보정(-)', '재고보정(+)'
]
ALL_PLANTS = ['화성', '대구', '부산', '대전']


# ── DB 연결 (읽기 전용) ───────────────────────────────────────────────────────
@st.cache_resource
def get_connection():
    if not os.path.exists(DB_PATH):
        st.error(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")
        st.stop()
    return duckdb.connect(DB_PATH, read_only=True)


con = get_connection()


# ── 공통 유틸 ─────────────────────────────────────────────────────────────────
@st.cache_data
def get_master_lists():
    partners = con.execute(
        "SELECT DISTINCT Partner_Name FROM Movement_Master ORDER BY Partner_Name"
    ).df()['Partner_Name'].tolist()
    products = con.execute(
        "SELECT DISTINCT Product_Name FROM Movement_Master ORDER BY Product_Name"
    ).df()['Product_Name'].tolist()
    return partners, products


def date_filter(all_date: bool, date_range) -> tuple[str, list]:
    """날짜 필터 SQL 조각과 파라미터 반환 - 파라미터 바인딩으로 SQL Injection 방지"""
    if all_date or len(date_range) != 2:
        return "1=1", []
    return "Date BETWEEN ? AND ?", [date_range[0], date_range[1]]


def in_clause(col: str, values: list) -> str:
    """IN 절 안전 생성 - 값을 직접 이스케이프하여 삽입 (숫자/문자열 모두 처리)"""
    escaped = ", ".join(f"'{str(v).replace(chr(39), chr(39)*2)}'" for v in values)
    return f"{col} IN ({escaped})"


def safe_query(query: str, params: list = []) -> pd.DataFrame:
    """쿼리 실행 + 오류 시 빈 DataFrame 반환"""
    try:
        return con.execute(query, params).df()
    except Exception as e:
        logging.error("쿼리 오류: %s | params: %s | err: %s", query[:80], params, e)
        st.error(f"데이터 조회 오류: {e}")
        return pd.DataFrame()


# ── 탭 구성 ───────────────────────────────────────────────────────────────────
all_partners, all_products = get_master_lists()
tab1, tab2, tab3, tab5, tab6 = st.tabs(
    ["📈 지점별 지표", "🔍 상세 검색", "📦 제품 / 🤝 매출처 순위", "🚀 영업 기회 예측", "📋 출하 내역 조회"]
)


# ─────────────────────────────────────────────────────────────────────────────
# [Tab 1] 지점별 지표 분석
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.title("📊 지점별 유통 분석 현황")

    with st.container():
        f1, f2, f3 = st.columns(3)
        with f1:
            date_range_t1 = st.date_input("조회 기간", [], key="t1_date")
        with f2:
            sel_descs = st.multiselect("분류 선택", ALL_MOVE_DESCS, default=['매입', '일반 출하'], key="t1_desc")
        with f3:
            sel_plants = st.multiselect("지점 선택", ALL_PLANTS, default=ALL_PLANTS, key="t1_plant")

    if sel_descs and sel_plants and len(date_range_t1) == 2:
        d_sql, d_params = date_filter(False, date_range_t1)

        query = f"""
            SELECT
                Date, Plant_Code,
                SUM(Amount) as Total_Amount,
                SUM(Qty)    as Total_Qty,
                COUNT(DISTINCT Ref_No) as Invoice_Count,
                strftime(Date, '%Y-%m') as YearMonth
            FROM Movement_Master
            WHERE {in_clause('Move_Desc', sel_descs)}
              AND {in_clause('Plant_Code', sel_plants)}
              AND {d_sql}
            GROUP BY Date, Plant_Code
            ORDER BY Date
        """
        with st.spinner('데이터 분석 중...'):
            df = safe_query(query, d_params)

        if not df.empty:
            st.divider()
            g1, g2 = st.columns(2)
            with g1:
                st.subheader("💰 지점별 금액(Amount) 추이")
                fig_amt = px.line(df, x='Date', y='Total_Amount', color='Plant_Code',
                                  labels={'Total_Amount': '금액 합계'}, template="plotly_white")
                st.plotly_chart(fig_amt, use_container_width=True)
            with g2:
                st.subheader("📦 지점별 수량(Qty) 추이")
                fig_qty = px.line(df, x='Date', y='Total_Qty', color='Plant_Code',
                                  labels={'Total_Qty': '수량 합계'}, template="plotly_white")
                st.plotly_chart(fig_qty, use_container_width=True)

            st.divider()
            st.subheader("📂 월별 상세 집계 내역")
            pivot = df.groupby(['YearMonth', 'Plant_Code']).agg(
                Total_Amount=('Total_Amount', 'sum'),
                Total_Qty=('Total_Qty', 'sum'),
                Invoice_Count=('Invoice_Count', 'sum')
            ).reset_index()

            t1, t2, t3_col = st.columns(3)
            with t1:
                st.markdown("**[월별 금액 집계 (원)]**")
                st.dataframe(
                    pivot.pivot(index='YearMonth', columns='Plant_Code', values='Total_Amount')
                         .fillna(0).style.format("{:,.0f}"),
                    use_container_width=True
                )
            with t2:
                st.markdown("**[월별 수량 집계]**")
                st.dataframe(
                    pivot.pivot(index='YearMonth', columns='Plant_Code', values='Total_Qty')
                         .fillna(0).style.format("{:,.0f}"),
                    use_container_width=True
                )
            with t3_col:
                st.markdown("**[월별 명세서 수 (건)]**")
                st.dataframe(
                    pivot.pivot(index='YearMonth', columns='Plant_Code', values='Invoice_Count')
                         .fillna(0).style.format("{:,.0f}"),
                    use_container_width=True
                )
        else:
            st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    else:
        st.info("조회 기간, 분류, 지점을 모두 선택해 주세요.")


# ─────────────────────────────────────────────────────────────────────────────
# [Tab 2] 상세 데이터 탐색
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.title("🔍 상세 데이터 탐색")

    with st.expander("검색 조건 설정", expanded=True):
        st.markdown("**📅 기간 설정**")
        d1, d2 = st.columns([1, 2])
        with d1:
            all_date_s = st.checkbox("전체 기간 검색", value=False, key="s_all_date")
        with d2:
            search_date = st.date_input("조회 기간 선택", [], disabled=all_date_s, key="s_date")

        st.markdown("**📍 지점 및 분류 필터**")
        fc1, fc2 = st.columns(2)
        with fc1:
            s_plants = st.multiselect("지점 선택", ALL_PLANTS, default=ALL_PLANTS, key="s_plant")
        with fc2:
            s_descs = st.multiselect("유통 분류 선택", ALL_MOVE_DESCS, default=['일반 출하'], key="s_desc")

        st.divider()
        st.markdown("**🔎 매출처 및 제품 정보**")
        r1, r2 = st.columns(2)
        with r1:
            sel_partner = st.selectbox("매출처명", ["전체조회"] + all_partners, key="s_partner")
        with r2:
            search_ref = st.text_input("명세서번호(Ref_No)", key="s_ref")
        r3, r4 = st.columns(2)
        with r3:
            sel_product = st.selectbox("제품명", ["전체조회"] + all_products, key="s_product")
        with r4:
            search_batch = st.text_input("제조번호(Batch)", key="s_batch")

    if st.button("🚀 데이터 검색 시작"):
        d_sql, params = date_filter(all_date_s, search_date)

        # 고정 조건 (파라미터 바인딩)
        conditions = [d_sql]

        # IN 절 (in_clause로 안전 처리)
        if s_plants:
            conditions.append(in_clause('Plant_Code', s_plants))
        if s_descs:
            conditions.append(in_clause('Move_Desc', s_descs))

        # LIKE / = 조건 (파라미터 바인딩)
        if search_batch:
            conditions.append("Batch LIKE ?")
            params.append(f"%{search_batch}%")
        if search_ref:
            conditions.append("Ref_No LIKE ?")
            params.append(f"%{search_ref}%")
        if sel_partner != "전체조회":
            conditions.append("Partner_Name = ?")
            params.append(sel_partner)
        if sel_product != "전체조회":
            conditions.append("Product_Name = ?")
            params.append(sel_product)

        where = " AND ".join(conditions)
        query = f"""
            SELECT Date, Time, Plant_Code, Move_Desc, Storage_Loc,
                   Partner_Name, Partner_Address,
                   Product_Code, Product_Name,
                   Batch, Qty, Amount, Ref_No
            FROM Movement_Master
            WHERE {where}
            ORDER BY Date DESC, Time DESC
            LIMIT 5000
        """
        with st.spinner('데이터를 불러오는 중...'):
            df_s = safe_query(query, params)

        if not df_s.empty:
            df_s['Date'] = pd.to_datetime(df_s['Date']).dt.date
            st.success(f"조회 성공: {len(df_s):,}건")
            st.dataframe(
                df_s.style.format({'Qty': '{:,.0f}', 'Amount': '{:,.0f}'}),
                use_container_width=True
            )
            csv = df_s.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                "💾 결과 저장 (CSV)", csv,
                f"search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", "text/csv"
            )
        else:
            st.warning("일치하는 데이터가 없습니다.")


# ─────────────────────────────────────────────────────────────────────────────
# [Tab 3] 제품 TOP 30 (왼쪽) + 거래처 TOP 10 (오른쪽)
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    st.title("📦 제품 / 🤝 매출처 순위 분석")

    # ── 공통 필터 ──
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        t3_plant = st.selectbox("분석 대상 지점", ALL_PLANTS, key="t3_plant")
    with fc2:
        t3_all = st.checkbox("전체 기간 기준", value=True, key="t3_all")
    with fc3:
        t3_date = st.date_input("조회 기간", [], disabled=t3_all, key="t3_date")

    d_sql, d_params = date_filter(t3_all, t3_date)
    plant_cond = in_clause('Plant_Code', [t3_plant])

    st.divider()

    # ── 좌우 레이아웃 ──
    left_col, right_col = st.columns(2)

    # ── 왼쪽: 제품 TOP 30 ──
    with left_col:
        for label, move_desc, icon in [
            ("일반 출하", "일반 출하", "🚀"),
            ("매입",     "매입",     "📥"),
        ]:
            st.subheader(f"{icon} 제품 {label} TOP 30")
            q = f"""
                SELECT Product_Code, Product_Name,
                       SUM(Qty) as Total_Qty, SUM(Amount) as Total_Amount
                FROM Movement_Master
                WHERE Move_Desc = ? AND {plant_cond} AND {d_sql}
                GROUP BY Product_Code, Product_Name
                ORDER BY Total_Amount DESC
                LIMIT 30
            """
            df_t3 = safe_query(q, [move_desc] + d_params)
            if not df_t3.empty:
                st.dataframe(
                    df_t3.style.format({'Total_Qty': '{:,.0f}', 'Total_Amount': '{:,.0f}'}),
                    use_container_width=True
                )
            else:
                st.warning("데이터가 없습니다.")
            st.divider()

    # ── 오른쪽: 거래처 TOP 10 ──
    with right_col:
        for label, move_desc, icon in [
            ("매출 높은 매출처 TOP 10", "일반 출하", "💰"),
            ("반품 금액 높은 매출처 TOP 10",  "일반 반품", "⚠️"),
        ]:
            st.subheader(f"{icon} {label}")
            q = f"""
                SELECT Partner_Name,
                       SUM(Amount) as Total_Amount,
                       SUM(Qty)    as Total_Qty
                FROM Movement_Master
                WHERE Move_Desc = ? AND {plant_cond} AND {d_sql}
                GROUP BY Partner_Name
                ORDER BY Total_Amount DESC
                LIMIT 10
            """
            df_t4 = safe_query(q, [move_desc] + d_params)
            if not df_t4.empty:
                st.dataframe(
                    df_t4.style.format({'Total_Qty': '{:,.0f}', 'Total_Amount': '{:,.0f}'}),
                    use_container_width=True
                )
            else:
                st.warning("데이터가 없습니다.")
            st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# [Tab 5] 영업 기회 예측
# ─────────────────────────────────────────────────────────────────────────────
with tab5:
    today      = datetime.now().date()
    this_year  = today.year
    this_month = today.month

    st.header("🚀 AI 영업 비서")

    c1, c2 = st.columns(2)
    with c1:
        lookback = st.slider("분석 기간(개월)", 1, 6, 3, help="최근 몇 개월의 평균을 계산할지 정합니다.")
    with c2:
        use_last_year = st.checkbox(
            "작년 동월 실적 반영 (하이브리드 분석)", value=False,
            help="체크 시 작년 같은 달 실적을 60% 비중으로 반영합니다."
        )

    st.caption(f"기준일: {today} | 분석 모델: 하이브리드 수요 예측 (최근 {lookback}개월 + 작년 동월)")

    start_calc = (today.replace(day=1) - timedelta(days=lookback * 30)).replace(day=1)
    end_calc   = today.replace(day=1) - timedelta(days=1)

    st.write(f"🔍 **분석 범위:** 작년 {this_month}월 실적 + 최근 {lookback}개월({start_calc} ~ {end_calc}) 평균")

    df_partners = safe_query("SELECT DISTINCT Partner_Name FROM Movement_Master ORDER BY Partner_Name")
    if df_partners.empty:
        st.warning("매출처 데이터가 없습니다.")
        st.stop()

    sel_pharmacy = st.selectbox("분석할 매출처를 선택하세요", df_partners['Partner_Name'], key="t5_pharmacy")

    if sel_pharmacy:
        query = """
        WITH recent_avg_data AS (
            SELECT Product_Name, AVG(Qty) as avg_recent
            FROM Movement_Master
            WHERE Partner_Name = ?
              AND Move_Desc = '일반 출하'
              AND Date >= ? AND Date <= ?
            GROUP BY Product_Name
        ),
        last_year_same_month AS (
            SELECT Product_Name, SUM(Qty) as last_year_qty
            FROM Movement_Master
            WHERE Partner_Name = ?
              AND Move_Desc = '일반 출하'
              AND YEAR(Date) = ? AND MONTH(Date) = ?
            GROUP BY Product_Name
        ),
        current_month AS (
            SELECT Product_Name, SUM(Qty) as current_qty
            FROM Movement_Master
            WHERE Partner_Name = ?
              AND Move_Desc = '일반 출하'
              AND YEAR(Date) = ? AND MONTH(Date) = ?
            GROUP BY Product_Name
        ),
        price_master AS (
            SELECT Product_Name, AVG(Amount / NULLIF(Qty, 0)) as avg_unit_price
            FROM Movement_Master
            WHERE Partner_Name = ? AND Move_Desc = '일반 출하' AND Qty > 0
            GROUP BY Product_Name
        )
        SELECT
            COALESCE(r.Product_Name, l.Product_Name, c.Product_Name) as "제품명",
            ROUND(COALESCE(r.avg_recent,    0), 1) as "최근 평균",
            ROUND(COALESCE(l.last_year_qty, 0), 1) as "작년 동월 실적",
            ROUND(COALESCE(c.current_qty,   0), 1) as "이번 달 현재 실적",
            ROUND(COALESCE(p.avg_unit_price,0), 0) as "최근 단가"
        FROM recent_avg_data r
        FULL OUTER JOIN last_year_same_month l ON r.Product_Name = l.Product_Name
        FULL OUTER JOIN current_month        c ON COALESCE(r.Product_Name, l.Product_Name) = c.Product_Name
        LEFT  JOIN price_master              p ON COALESCE(r.Product_Name, l.Product_Name, c.Product_Name) = p.Product_Name
        """
        params = [
            sel_pharmacy, str(start_calc), str(end_calc),   # recent_avg_data
            sel_pharmacy, this_year - 1, this_month,        # last_year_same_month
            sel_pharmacy, this_year, this_month,            # current_month
            sel_pharmacy,                                   # price_master
        ]

        df_fc = safe_query(query, params)

        if df_fc.empty:
            st.warning("해당 매출처의 데이터가 없습니다.")
        else:
            # 예측 수요 계산
            def calc_forecast(row):
                if not use_last_year or row["작년 동월 실적"] == 0:
                    return row["최근 평균"]
                return row["최근 평균"] * 0.4 + row["작년 동월 실적"] * 0.6

            df_fc["예측 수요"]       = df_fc.apply(calc_forecast, axis=1).round(1)
            df_fc["추가 제안 가능량"] = (df_fc["예측 수요"] - df_fc["이번 달 현재 실적"]).clip(lower=0)
            df_fc["예상 추가 매출"]   = (df_fc["추가 제안 가능량"] * df_fc["최근 단가"]).astype(int)

            df_result = (
                df_fc[df_fc["예측 수요"] > 0]
                .sort_values("예상 추가 매출", ascending=False)
                .reset_index(drop=True)
            )

            def highlight_top30(row):
                # 들여쓰기 수정 (기존 IndentationError 수정)
                if row.name < 30:
                    return ['background-color: #2E7D32; color: #FFFFFF; font-weight: bold;'] * len(row)
                return [''] * len(row)

            st.subheader(f"📍 {sel_pharmacy} 영업 기회 분석 결과")
            st.info("✅ 상위 30개 핵심 품목이 진한 녹색으로 강조 표시됩니다.")

            st.dataframe(
                df_result.style.apply(highlight_top30, axis=1).format({
                    "최근 평균": "{:.1f}", "작년 동월 실적": "{:.1f}",
                    "이번 달 현재 실적": "{:.1f}", "예측 수요": "{:.1f}",
                    "추가 제안 가능량": "{:.1f}", "최근 단가": "{:,.0f}", "예상 추가 매출": "{:,.0f}"
                }),
                use_container_width=True
            )

            # 엑셀 다운로드
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, sheet_name='Sales_Opportunity')

            st.download_button(
                label="📊 이 분석 결과 엑셀로 내려받기",
                data=output.getvalue(),
                file_name=f"{sel_pharmacy}_sales_proposal_{today}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ─────────────────────────────────────────────────────────────────────────────
# [Tab 6] 일반 출하 내역 조회 (지점 + 날짜 + 시간 범위)
# ─────────────────────────────────────────────────────────────────────────────
with tab6:
    st.title("📋 일반 출하 내역 조회")
    st.caption("지점, 시작 날짜/시간, 종료 날짜/시간을 선택하면 일반 출하 데이터를 조회하고 엑셀로 저장할 수 있습니다.")

    st.divider()

    # ── 필터 영역 ──
    st.markdown("**지점 선택**")
    t6_plant = st.selectbox("지점", ALL_PLANTS, key="t6_plant", label_visibility="collapsed")

    st.markdown("**시작 날짜 / 시간**")
    sc1, sc2 = st.columns(2)
    with sc1:
        t6_date_start = st.date_input("시작 날짜", key="t6_date_start")
    with sc2:
        t6_time_start = st.time_input("시작 시간", value=None, key="t6_time_start", step=3600)

    st.markdown("**종료 날짜 / 시간**")
    ec1, ec2 = st.columns(2)
    with ec1:
        t6_date_end = st.date_input("종료 날짜", key="t6_date_end")
    with ec2:
        t6_time_end = st.time_input("종료 시간", value=None, key="t6_time_end", step=3600)

    st.divider()

    if st.button("🔍 조회", key="t6_search"):
        # 유효성 검사
        if t6_time_start is None or t6_time_end is None:
            st.warning("시작 시간과 종료 시간을 모두 입력해 주세요.")
            st.stop()

        from datetime import datetime as dt
        dt_start = dt.combine(t6_date_start, t6_time_start)
        dt_end   = dt.combine(t6_date_end,   t6_time_end)

        if dt_start >= dt_end:
            st.warning(f"시작({dt_start.strftime('%Y-%m-%d %H:%M')})이 종료({dt_end.strftime('%Y-%m-%d %H:%M')})보다 앞이어야 합니다.")
            st.stop()

        plant_cond_t6 = in_clause('Plant_Code', [t6_plant])

        query = f"""
            SELECT
                Date, Time, Plant_Code,
                Partner_Name, Partner_Address,
                Product_Code, Product_Name,
                Batch, Qty, Amount, Ref_No
            FROM Movement_Master
            WHERE Move_Desc = '일반 출하'
              AND {plant_cond_t6}
              AND (Date > ? OR (Date = ? AND Time >= ?))
              AND (Date < ? OR (Date = ? AND Time <= ?))
            ORDER BY Date ASC, Time ASC
        """
        params = [
            str(t6_date_start), str(t6_date_start), str(t6_time_start),
            str(t6_date_end),   str(t6_date_end),   str(t6_time_end)
        ]

        with st.spinner("데이터 조회 중..."):
            df_t6 = safe_query(query, params)

        if df_t6.empty:
            st.warning("조건에 맞는 데이터가 없습니다.")
        else:
            df_t6['Date'] = pd.to_datetime(df_t6['Date']).dt.date
            st.success(f"조회 결과: {len(df_t6):,}건  |  {dt_start.strftime('%Y-%m-%d %H:%M')} ~ {dt_end.strftime('%Y-%m-%d %H:%M')}")
            st.dataframe(
                df_t6.style.format({'Qty': '{:,.0f}', 'Amount': '{:,.0f}'}),
                use_container_width=True
            )

            # 엑셀 다운로드
            output_t6 = io.BytesIO()
            with pd.ExcelWriter(output_t6, engine='xlsxwriter') as writer:
                df_t6.to_excel(writer, index=False, sheet_name='출하내역')

            st.download_button(
                label="📥 엑셀로 내려받기",
                data=output_t6.getvalue(),
                file_name=f"{t6_plant}_출하내역_{dt_start.strftime('%Y%m%d_%H%M')}_{dt_end.strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )