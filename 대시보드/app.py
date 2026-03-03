import streamlit as st
import pandas as pd
import duckdb
import plotly.express as px  # <-- 이 줄을 꼭 추가해야 합니다!
# timedelta를 명시적으로 추가해줍니다.
from datetime import datetime, date, timedelta
import io
import xlsxwriter     # <- 이 줄도 확인 (없으면 에러 날 수 있음)
import os  # <-- 이 줄이 없어서 에러가 발생한 것입니다!

# 1. 페이지 및 경로 설정
st.set_page_config(page_title="의약품 유통 분석 시스템 v1.3", layout="wide")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'pharmacy_data.db')

@st.cache_resource
def get_connection():
    if not os.path.exists(DB_PATH):
        st.error(f"DB 파일을 찾을 수 없습니다: {DB_PATH}")
        st.stop()
    return duckdb.connect(DB_PATH, read_only=True)

con = get_connection()

# [수정] 이제 코드가 아닌 '명칭' 리스트를 직접 정의합니다. (약가인하 포함)
ALL_MOVE_DESCS = [
    '입고', '일반 출하', '약가인하', '출하 취소', '이월 매출', 
    '지점 이동', '일반 반품', 'RTP', 'POST', '스크랩', '재고보정(-)', '재고보정(+)'
]

# 마스터 리스트 캐싱
@st.cache_data
def get_master_lists():
    partners = con.execute("SELECT DISTINCT Partner_Name FROM Movement_Master ORDER BY Partner_Name").df()['Partner_Name'].tolist()
    products = con.execute("SELECT DISTINCT Product_Name FROM Movement_Master ORDER BY Product_Name").df()['Product_Name'].tolist()
    return partners, products

all_partners, all_products = get_master_lists()

# 탭 구성
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 지점별 지표", "🔍 상세 검색", "📦 제품 TOP 30", "🤝 거래처 TOP 10", "🚀 영업 기회 예측"])

# ---------------------------------------------------------
# [공통 필터 기능] - 탭 내부에서 활용
# ---------------------------------------------------------
def get_date_filter_sql(all_date_search, search_date_range):
    if all_date_search:
        return "1=1", []
    if len(search_date_range) == 2:
        return "Date BETWEEN ? AND ?", [search_date_range[0], search_date_range[1]]
    return "1=1", []

# ---------------------------------------------------------
# [Tab 1] 지점별 지표 분석 (그래프 및 월별 집계표)(수정 완료 버전)
# ---------------------------------------------------------
# ---------------------------------------------------------
# [Tab 1] 지점별 지표 분석
# ---------------------------------------------------------
with tab1:
    st.title("📊 지점별 유통 분석 현황")
    
    # 이 아래줄의 시작 위치(들여쓰기)를 확인하세요!
    with st.container():
        f_col1, f_col2, f_col3 = st.columns([2, 2, 2])
        with f_col1:
            date_range = st.date_input("조회 기간", [], key="tab1_date_input")
        with f_col2:
            selected_descs = st.multiselect("분류 선택", options=ALL_MOVE_DESCS, default=['입고', '일반 출하'])
        with f_col3:
            selected_plants = st.multiselect("지점 선택", options=['화성', '대구', '부산', '대전'], default=['화성', '대구', '부산', '대전'])

    if selected_descs and selected_plants and len(date_range) == 2:
        # 1. [보완] 공통 함수를 사용하여 날짜 쿼리문과 파라미터를 가져옵니다.
        date_sql, date_params = get_date_filter_sql(False, date_range)
        
        # 2. [핵심] Move_Desc를 사용하여 '약가인하'가 자동으로 분리됩니다.
        query = f"""
            SELECT 
                Date, Plant_Code, SUM(Amount) as Total_Amount, SUM(Qty) as Total_Qty,
                COUNT(DISTINCT Ref_No) as Invoice_Count,
                strftime(Date, '%Y-%m') as YearMonth
            FROM Movement_Master
            WHERE Move_Desc IN {tuple(selected_descs) if len(selected_descs) > 1 else f"('{selected_descs[0]}')"}
              AND Plant_Code IN {tuple(selected_plants) if len(selected_plants) > 1 else f"('{selected_plants[0]}')"}
              AND {date_sql}
            GROUP BY Date, Plant_Code
            ORDER BY Date
        """
        
        # 3. [보완] 파라미터(date_params)를 함께 전달하여 실행합니다.
        with st.spinner('데이터 분석 중...'):
            df = con.execute(query, date_params).df()

        if not df.empty:
            # --- 그래프 및 표 출력 로직 (파트너님 코드와 동일) ---
            st.divider()
            g_col1, g_col2 = st.columns(2)
            
            with g_col1:
                st.subheader("💰 지점별 금액(Amount) 추이")
                fig_amt = px.line(df, x='Date', y='Total_Amount', color='Plant_Code', 
                                  labels={'Total_Amount': '금액 합계'}, template="plotly_white")
                st.plotly_chart(fig_amt, use_container_width=True)

            with g_col2:
                st.subheader("📦 지점별 수량(Qty) 추이")
                fig_qty = px.line(df, x='Date', y='Total_Qty', color='Plant_Code', 
                                  labels={'Total_Qty': '수량 합계'}, template="plotly_white")
                st.plotly_chart(fig_qty, use_container_width=True)

            st.divider()
            st.subheader("📂 월별 상세 집계 내역")
            
            pivot_df = df.groupby(['YearMonth', 'Plant_Code']).agg({
                'Total_Amount': 'sum', 'Total_Qty': 'sum', 'Invoice_Count': 'sum'
            }).reset_index()

            t_col1, t_col2, t_col3 = st.columns(3)
            with t_col1:
                st.markdown("**[월별 금액 집계 (원)]**")
                amt_table = pivot_df.pivot(index='YearMonth', columns='Plant_Code', values='Total_Amount').fillna(0)
                st.dataframe(amt_table.style.format("{:,.0f}"), use_container_width=True)
            with t_col2:
                st.markdown("**[월별 수량 집계]**")
                qty_table = pivot_df.pivot(index='YearMonth', columns='Plant_Code', values='Total_Qty').fillna(0)
                st.dataframe(qty_table.style.format("{:,.0f}"), use_container_width=True)
            with t_col3:
                st.markdown("**[월별 명세서 수 (건)]**")
                inv_table = pivot_df.pivot(index='YearMonth', columns='Plant_Code', values='Invoice_Count').fillna(0)
                st.dataframe(inv_table.style.format("{:,.0f}"), use_container_width=True)
        else:
            st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    else:
        st.info("조회 기간, 분류, 지점을 모두 선택해 주세요.")

# ---------------------------------------------------------
# [Tab 2] 상세 데이터 탐색 (보관장소 Storage_Loc 포함)
# ---------------------------------------------------------
with tab2:
    st.title("🔍 상세 데이터 탐색")
    
    with st.expander("검색 조건 설정", expanded=True):
        st.markdown("**📅 기간 설정**")
        d_col1, d_col2 = st.columns([1, 2])
        with d_col1:
            all_date_search = st.checkbox("전체 기간 검색", value=False)
        with d_col2:
            search_date_range = st.date_input("조회 기간 선택", [], disabled=all_date_search)

        st.divider()

        st.markdown("**🔎 상세 정보 입력**")
        s_col1, s_col2, s_col3 = st.columns([1, 1, 1])
        with s_col1:
            search_batch = st.text_input("제조번호(Batch)")
            search_ref = st.text_input("명세서번호(Ref_No)")
        with s_col2:
            # 이제 DB에 '미등록(코드)'으로 저장되어 있어 검색이 훨씬 정확해집니다.
            selected_partner = st.selectbox("거래처명", options=["전체조회"] + all_partners)
        with s_col3:
            selected_product = st.selectbox("제품명", options=["전체조회"] + all_products)

    if st.button("🚀 데이터 검색 시작"):
        # 1. 기본 쿼리 (Storage_Loc 추가 및 컬럼 순서 조정)
        detail_query = """
            SELECT 
                Date, Time, Plant_Code, 
                Move_Desc, Storage_Loc,         -- [추가] 이제 분류명과 보관장소를 같이 봅니다.
                Partner_Name, Partner_Address, 
                Product_Name, Batch, Qty, Amount, Ref_No 
            FROM Movement_Master 
            WHERE 1=1
        """
        
        # 2. 날짜 필터 처리 (공통 함수 사용)
        date_sql, date_params = get_date_filter_sql(all_date_search, search_date_range)
        detail_query += f" AND {date_sql}"
        params = date_params # 날짜 파라미터로 시작

        # 3. 추가 필터 (파라미터 방식 유지 - 보안상 매우 중요!)
        if search_batch:
            detail_query += " AND Batch LIKE ?"
            params.append(f"%{search_batch}%")
        if search_ref:
            detail_query += " AND Ref_No LIKE ?"
            params.append(f"%{search_ref}%")
        if selected_partner != "전체조회":
            detail_query += " AND Partner_Name = ?"
            params.append(selected_partner)
        if selected_product != "전체조회":
            detail_query += " AND Product_Name = ?"
            params.append(selected_product)

        # 4. 정렬 및 실행
        detail_query += " ORDER BY Date DESC, Time DESC LIMIT 5000"

        with st.spinner('데이터를 불러오는 중...'):
            try:
                # [중요] 쿼리와 params를 함께 전달합니다.
                search_results = con.execute(detail_query, params).df()

                if not search_results.empty:
                    st.success(f"조회 성공: {len(search_results):,}건")
                    st.dataframe(
                        search_results.style.format({'Qty': '{:,.0f}', 'Amount': '{:,.0f}'}), 
                        use_container_width=True
                    )
                    
                    csv = search_results.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("💾 결과 저장 (CSV)", csv, f"search_{datetime.now().strftime('%H%M%S')}.csv", "text/csv")
                else:
                    st.warning("데이터가 없습니다.")
            except Exception as e:
                st.error(f"오류 발생: {e}")

# ---------------------------------------------------------
# [Tab 3] 제품 분석 (출하/입고 TOP 30)
# ---------------------------------------------------------
# [Tab 3] 수정 코드 부분
with tab3:
    st.title("📦 제품별 유통 순위 (TOP 30)")
    
    t3_col1, t3_col2 = st.columns(2)
    with t3_col1:
        target_plant = st.selectbox("분석 대상 지점", options=['화성', '대구', '부산', '대전'], key="t3_plant")
    with t3_col2:
        t3_all_date = st.checkbox("전체 기간 기준", value=True, key="t3_all_date")
        t3_date_range = st.date_input("조회 기간", [], disabled=t3_all_date, key="t3_date")

    date_sql, date_params = get_date_filter_sql(t3_all_date, t3_date_range)

    # --- [출하 및 입고 통합 표] ---
    st.divider()
    
    # 601: 일반 출하
    st.subheader(f"🚀 {target_plant} 지점 - 출하(601) 통합 TOP 30")
    df_ship_total = con.execute(f"""
        SELECT Product_Name, SUM(Qty) as Total_Qty, SUM(Amount) as Total_Amount
        FROM Movement_Master WHERE Move_Type = '601' AND Plant_Code = '{target_plant}' AND {date_sql}
        GROUP BY Product_Name ORDER BY Total_Amount DESC LIMIT 30
    """, date_params).df()
    st.dataframe(df_ship_total.style.format({'Total_Qty': '{:,.0f}', 'Total_Amount': '{:,.0f}'}), use_container_width=True)

    st.divider()

    # 101: 입고
    st.subheader(f"📥 {target_plant} 지점 - 입고(101) 통합 TOP 30")
    df_rec_total = con.execute(f"""
        SELECT Product_Name, SUM(Qty) as Total_Qty, SUM(Amount) as Total_Amount
        FROM Movement_Master WHERE Move_Type = '101' AND Plant_Code = '{target_plant}' AND {date_sql}
        GROUP BY Product_Name ORDER BY Total_Amount DESC LIMIT 30
    """, date_params).df()
    st.dataframe(df_rec_total.style.format({'Total_Qty': '{:,.0f}', 'Total_Amount': '{:,.0f}'}), use_container_width=True)

# ---------------------------------------------------------
# [Tab 4] 거래처 분석 (매출/반품 TOP 10)
# ---------------------------------------------------------
with tab4:
    st.title("🤝 거래처별 분석 (TOP 10)")
    
    t4_col1, t4_col2 = st.columns(2)
    with t4_col1:
        target_plant_p = st.selectbox("분석 대상 지점", options=['화성', '대구', '부산', '대전'], key="t4_plant")
    with t4_col2:
        t4_all_date = st.checkbox("전체 기간 기준", value=True, key="t4_all_date")
        t4_date_range = st.date_input("조회 기간", [], disabled=t4_all_date, key="t4_date")

    date_sql_p, date_params_p = get_date_filter_sql(t4_all_date, t4_date_range)

    st.divider()
    pc1, pc2 = st.columns(2)
    
    with pc1:
        st.subheader(f"💰 매출(출하) 높은 약국 TOP 10")
        df_top_sales = con.execute(f"""
            SELECT Partner_Name, SUM(Amount) as Total_Amount, SUM(Qty) as Total_Qty
            FROM Movement_Master WHERE Move_Type = '601' AND Plant_Code = '{target_plant_p}' AND {date_sql_p}
            GROUP BY Partner_Name ORDER BY Total_Amount DESC LIMIT 10
        """, date_params_p).df()
        st.dataframe(df_top_sales.style.format({'Total_Qty': '{:,.0f}', 'Total_Amount': '{:,.0f}'}), use_container_width=True)

    with pc2:
        st.subheader(f"⚠️ 반품 금액 높은 약국 TOP 10")
        df_top_return = con.execute(f"""
            SELECT Partner_Name, SUM(Amount) as Total_Amount, SUM(Qty) as Total_Qty
            FROM Movement_Master WHERE Move_Type = '655' AND Plant_Code = '{target_plant_p}' AND {date_sql_p}
            GROUP BY Partner_Name ORDER BY Total_Amount DESC LIMIT 10
        """, date_params_p).df()
        st.dataframe(df_top_return.style.format({'Total_Qty': '{:,.0f}', 'Total_Amount': '{:,.0f}'}), use_container_width=True)

# ---------------------------------------------------------
# 탭 정의에 tab5 추가
# tab1, tab2, tab3, tab4, tab5 = st.tabs(["조회", "분석", "...", "...", "🚀 영업 기회 예측"])
# ---------------------------------------------------------

with tab5:
    # 1. 날짜 설정  (사이드바 대신 변수만 조용히 선언)
    today = datetime.now().date()
    this_year = today.year
    this_month = today.month

    st.header("🚀 AI 영업 비서")

    # 설정 영역
    c1, c2 = st.columns(2)
    with c1:
        lookback_period = st.slider("분석 기간(개월)", 1, 6, 3, help="최근 몇 개월의 평균을 계산할지 정합니다.")
    with c2:
        # [신규 추가] 작년 데이터 포함 여부 체크박스
        use_last_year = st.checkbox("작년 동월 실적 반영 (하이브리드 분석)", value=False, 
                                     help="체크 시 작년 같은 달 실적을 60% 비중으로 반영합니다.")
        
    # 3. 이제 변수가 생겼으니 캡션에서 마음껏 사용합니다
    st.caption(f"기준일: {today} | 분석 모델: 하이브리드 수요 예측 (최근 {lookback_period}개월 + 작년 동월)")

    # 4. 분석 시작 날짜 계산 (이후 로직은 동일)
    start_date_calc = (today.replace(day=1) - timedelta(days=lookback_period * 30)).replace(day=1)
    end_date_calc = today.replace(day=1) - timedelta(days=1) # 전월 말일까지

    st.write(f"🔍 **분석 범위:** 작년 {this_month}월 실적 + 최근 {lookback_period}개월({start_date_calc} ~ {end_date_calc}) 평균")

    # 3. 약국 선택
    all_partners = con.execute("SELECT DISTINCT Partner_Name FROM Movement_Master ORDER BY Partner_Name").df()
    selected_pharmacy = st.selectbox("분석할 약국을 선택하세요", all_partners['Partner_Name'], key="forecast_pharmacy")

    if selected_pharmacy:
        # 4. 동적 쿼리 (선택한 날짜와 기간 반영)
        query = f"""
        WITH recent_avg_data AS (
            SELECT Product_Name, AVG(Qty) as avg_recent
            FROM Movement_Master
            WHERE Partner_Name = '{selected_pharmacy}'
              AND Date >= '{start_date_calc}' AND Date <= '{end_date_calc}'
            GROUP BY Product_Name
        ),
        last_year_same_month AS (
            SELECT Product_Name, SUM(Qty) as last_year_qty
            FROM Movement_Master
            WHERE Partner_Name = '{selected_pharmacy}'
              AND YEAR(Date) = {this_year - 1} AND MONTH(Date) = {this_month}
            GROUP BY Product_Name
        ),
        current_month AS (
            SELECT Product_Name, SUM(Qty) as current_qty
            FROM Movement_Master
            WHERE Partner_Name = '{selected_pharmacy}'
              AND YEAR(Date) = {this_year} AND MONTH(Date) = {this_month}
            GROUP BY Product_Name
        ),
        price_master AS (
            SELECT Product_Name, AVG(Amount / NULLIF(Qty, 0)) as avg_unit_price
            FROM Movement_Master
            WHERE Partner_Name = '{selected_pharmacy}' AND Qty > 0
            GROUP BY Product_Name
        )
        SELECT 
            COALESCE(r.Product_Name, l.Product_Name, c.Product_Name) as "제품명",
            ROUND(COALESCE(r.avg_recent, 0), 1) as "최근 평균",
            ROUND(COALESCE(l.last_year_qty, 0), 1) as "작년 동월 실적",
            ROUND(COALESCE(c.current_qty, 0), 1) as "이번 달 현재 실적",
            ROUND(COALESCE(p.avg_unit_price, 0), 0) as "최근 단가"
        FROM recent_avg_data r
        FULL OUTER JOIN last_year_same_month l ON r.Product_Name = l.Product_Name
        FULL OUTER JOIN current_month c ON COALESCE(r.Product_Name, l.Product_Name) = c.Product_Name
        LEFT JOIN price_master p ON COALESCE(r.Product_Name, l.Product_Name, c.Product_Name) = p.Product_Name
        """

        df_forecast = con.execute(query).df()
        
        # [로직 변경] 작년 실적 반영 여부에 따른 예측 계산
        def calculate_forecast(row):
            if not use_last_year or row["작년 동월 실적"] == 0:
                return row["최근 평균"]  # 작년 데이터 무시하고 최근 평균만 100% 사용
            return (row["최근 평균"] * 0.4) + (row["작년 동월 실적"] * 0.6)

        df_forecast["예측 수요"] = df_forecast.apply(calculate_forecast, axis=1).round(1)
        df_forecast["추가 제안 가능량"] = (df_forecast["예측 수요"] - df_forecast["이번 달 현재 실적"]).clip(lower=0)
        df_forecast["예상 추가 매출"] = (df_forecast["추가 제안 가능량"] * df_forecast["최근 단가"]).astype(int)

        # 결과 정렬 (금액 높은 순)
        df_result = df_forecast[df_forecast["예측 수요"] > 0].sort_values(by="예상 추가 매출", ascending=False).reset_index(drop=True)

        # [신규 추가] 상위 30개 행 색칠하기 함수
        def highlight_top_30(row):
# row.name은 현재 데이터프레임의 인덱스입니다. (0~29번 행까지 색칠)
            if row.name < 30:
                # 배경: 진한 녹색 (#2E7D32), 글자: 흰색 (#FFFFFF), 글꼴: 굵게
                return ['background-color: #2E7D32; color: #FFFFFF; font-weight: bold;'] * len(row)
            return [''] * len(row)

        st.subheader(f"📍 {selected_pharmacy} 영업 기회 분석 결과")
        st.info("✅ 상위 30개 핵심 품목이 진한 녹색으로 강조 표시됩니다.")

        # 데이터프레임 스타일 적용 및 출력
        st.dataframe(
            df_result.style.apply(highlight_top_30, axis=1).format({
                "최근 평균": "{:.1f}", "작년 동월 실적": "{:.1f}", 
                "이번 달 현재 실적": "{:.1f}", "예측 수요": "{:.1f}",
                "추가 제안 가능량": "{:.1f}", "최근 단가": "{:,.0f}", "예상 추가 매출": "{:,.0f}"
            }), 
            use_container_width=True
        )

        # 6. 엑셀 다운로드 기능
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Sales_Opportunity')
        
        st.download_button(
            label="📊 이 분석 결과 엑셀로 내려받기",
            data=output.getvalue(),
            file_name=f"{selected_pharmacy}_영업제안_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
