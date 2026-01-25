import pandas as pd
import numpy as np
import folium
from geopy.distance import geodesic
import copy

# ==========================================
# [설정] 통합 수정판 + 엑셀 요약/유류비 추가
# ==========================================
TARGET_VEHICLE_COUNT = 11
AVG_SPEED_KMH = 35
SERVICE_TIME_MIN = 5

# [신규] 유류비 설정 (사용자가 변경 가능)
FUEL_PRICE = 1700  # 리터당 가격 (원)
DEFAULT_EFFICIENCY = 9.5 # 기본 연비 (km/L) - 1톤 트럭 기준

# 차량별 담당 키워드
ROUTE_ASSIGNMENTS = {
    0: ['해운대', '기장', '정관'],      # SB01
    1: ['반송', '반여', '재송', '수영'], # SB02
    2: ['울산', '정관'],              # SB03
    3: ['동구', '남구'],              # SB04
    4: ['김해', '창원', '진해', '마산'], # SB05
    5: ['동구', '진구'],              # SB06
    6: ['북구', '사상', '강서'],        # SB07
    7: ['양산', '서창', '금정'],        # SB08
    8: ['사하', '서구'],              # SB09
    9: ['금정', '동래', '연제'],        # SB10
    10: ['서구', '중구', '영도']        # SB11
}

# 강제 밸런싱 타겟
FORCE_PAIRS = [
    (8, 10), # SB09 -> SB11
    (3, 5),  # SB04 -> SB06
    (7, 9),  # SB08 -> SB10
    (6, 5),  # SB07 -> SB06
    (2, 0)   # SB03 -> SB01
]
# ==========================================

print("🔥 [최종 완료] 유류비 계산 및 요약 시트가 포함된 배차를 시작합니다...")

# 시간 및 거리 계산 함수
def calculate_metrics(start_pos, orders):
    if not orders: return 0, 0
    rem = orders[:]
    curr = start_pos
    total_dist = 0
    while rem:
        best_idx = min(range(len(rem)), key=lambda i: geodesic(curr, rem[i]['coord']).km)
        tgt = rem.pop(best_idx)
        total_dist += geodesic(curr, tgt['coord']).km
        curr = tgt['coord']
    total_dist += geodesic(curr, start_pos).km
    
    total_time = (total_dist/AVG_SPEED_KMH)*60 + len(orders)*SERVICE_TIME_MIN
    return total_dist, total_time

try:
    # 1. 데이터 로드 (연비 컬럼 추가 확인)
    df_vehicle = pd.read_excel("배차데이터_양식.xlsx", sheet_name="차량정보")
    
    def parse_coord(raw):
        try:
            if ',' in str(raw):
                parts = raw.split(',')
                return [float(parts[1]), float(parts[0])]
        except: return None
        return None

    vehicles = []
    starts = []
    
    # [신규] 연비 컬럼 감지
    eff_col = '연비'
    if '연비' not in df_vehicle.columns:
        print("⚠️ 차량정보에 '연비' 컬럼이 없어 기본값(9.5km/L)을 사용합니다.")
        eff_col = None

    for i, row in df_vehicle.iterrows():
        if i >= 11: break
        c = parse_coord(row.get('출발지'))
        if not c: c = [35.1795543, 129.0756416]
        starts.append(c)
        
        # 연비 가져오기
        efficiency = row.get(eff_col, DEFAULT_EFFICIENCY) if eff_col else DEFAULT_EFFICIENCY
        
        vehicles.append({
            'idx': i, 'name': row['차량명'], 'start': c,
            'keywords': ROUTE_ASSIGNMENTS.get(i, []),
            'efficiency': float(efficiency), # 연비 저장
            'orders': [], 'time': 0, 'dist': 0 # 거리 변수 추가
        })

    depot_center = np.mean(starts, axis=0).tolist()
    df_order = pd.read_excel("배차데이터_좌표완료.xlsx", sheet_name="주문목록").head(240)
    
    # 이름 컬럼 자동 감지
    name_col = None
    possible_names = ['name', 'Name', '상호명', '거래처명', '장소명', '이름', '고객명']
    for col in df_order.columns:
        if col in possible_names: name_col = col; break

    all_orders = []
    for i, row in df_order.iterrows():
        c = parse_coord(row['좌표'])
        if c:
            addr = str(row['address'])
            nm = str(row[name_col]) if name_col else f"주문_{i+1}"
            cands = [v['idx'] for v in vehicles if any(k in addr for k in v['keywords'])]
            all_orders.append({
                'id': i, 'name': nm, 'coord': c, 'addr': addr,
                'candidates': cands, 'assigned': False
            })

    # 2. 초기 배정
    print("Step 1: 1차 배정...")
    for o in all_orders: # 단독
        if len(o['candidates']) == 1:
            v = vehicles[o['candidates'][0]]
            v['orders'].append(o); o['assigned'] = True
    for o in [x for x in all_orders if not x['assigned'] and len(x['candidates']) > 1]: # 중복
        best_v = min([vehicles[i] for i in o['candidates']], key=lambda v: geodesic(v['start'], o['coord']).km)
        best_v['orders'].append(o); o['assigned'] = True
    for o in [x for x in all_orders if not x['assigned']]: # 낙오자
        best_v = min(vehicles, key=lambda v: geodesic(v['start'], o['coord']).km)
        best_v['orders'].append(o); o['assigned'] = True

    for v in vehicles: 
        d, t = calculate_metrics(v['start'], v['orders'])
        v['dist'] = d; v['time'] = t

    # 3. 강제 밸런싱
    print("Step 2: 강제 평준화...")
    for busy_idx, idle_idx in FORCE_PAIRS:
        busy, idle = vehicles[busy_idx], vehicles[idle_idx]
        for _ in range(100):
            d1, t1 = calculate_metrics(busy['start'], busy['orders'])
            d2, t2 = calculate_metrics(idle['start'], idle['orders'])
            busy['time'] = t1; idle['time'] = t2
            
            if t1 - t2 < 20: break
            
            # 거리 제한 없이 이동
            best_idx = -1; min_dist = float('inf')
            for idx, o in enumerate(busy['orders']):
                d = geodesic(idle['start'], o['coord']).km
                if d < min_dist: min_dist = d; best_idx = idx
            
            if best_idx != -1:
                idle['orders'].append(busy['orders'].pop(best_idx))
            else: break

    # 4. 전체 미세 조정
    print("Step 3: 전체 미세 조정...")
    for _ in range(10):
        vehicles.sort(key=lambda x: x['time'], reverse=True)
        b, i = vehicles[0], vehicles[-1]
        if b['time'] - i['time'] < 30: break
        
        best_idx = -1; min_d = 15.0
        for idx, o in enumerate(b['orders']):
            d = geodesic(i['start'], o['coord']).km
            if d < min_d: min_d = d; best_idx = idx
        if best_idx != -1:
            i['orders'].append(b['orders'].pop(best_idx))
            d1, t1 = calculate_metrics(b['start'], b['orders'])
            d2, t2 = calculate_metrics(i['start'], i['orders'])
            b['time'] = t1; i['time'] = t2

    # 재정렬 및 최종 계산
    vehicles.sort(key=lambda x: x['idx'])
    
    # -----------------------------------------------------------
    # [신규] 5. 결과 및 요약 시트 생성
    # -----------------------------------------------------------
    print("🚀 엑셀 결과 및 요약 생성 중...")
    
    # (1) 상세 목록용 DataFrame
    df_detail = pd.read_excel("배차데이터_좌표완료.xlsx", sheet_name="주문목록").head(240)
    df_detail['배차차량'] = "미배차"; df_detail['예상도착시간'] = ""
    
    # (2) 요약용 데이터 리스트
    summary_list = []
    
    m = folium.Map(location=depot_center, zoom_start=11)
    colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue', 'black', 'gray', 'darkblue', 'darkgreen']
    
    for v in vehicles:
        c = colors[v['idx'] % len(colors)]
        rem = v['orders'][:]; curr = v['start']; tm = 0; txt = []; path = [v['start']]
        
        # TSP 경로 확정 및 상세 엑셀 기록
        while rem:
            idx = min(range(len(rem)), key=lambda i: geodesic(curr, rem[i]['coord']).km)
            tgt = rem.pop(idx)
            tm += (geodesic(curr, tgt['coord']).km / AVG_SPEED_KMH)*60 + SERVICE_TIME_MIN
            curr = tgt['coord']
            path.append(curr); txt.append(tgt['name'])
            arr = int(tm - SERVICE_TIME_MIN)
            df_detail.at[tgt['id'], '배차차량'] = v['name']
            df_detail.at[tgt['id'], '예상도착시간'] = f"{9+arr//60:02d}:{arr%60:02d}"
        
        path.append(v['start'])
        return_dist = geodesic(curr, v['start']).km
        tm += (return_dist/AVG_SPEED_KMH)*60
        
        # [신규] 총 거리 및 유류비 계산
        total_dist_km, _ = calculate_metrics(v['start'], v['orders'])
        fuel_cost = int((total_dist_km / v['efficiency']) * FUEL_PRICE)
        
        # 요약 데이터 추가
        summary_list.append({
            '차량명': v['name'],
            '배송지수': len(v['orders']),
            '총거리(km)': round(total_dist_km, 1),
            '총시간(h)': round(tm/60, 1),
            '연비(km/L)': v['efficiency'],
            '예상유류비(원)': f"{fuel_cost:,}"
        })
        
        kw = ",".join(v['keywords'][:2])
        print(f"[{v['name']}] {kw}.. | {len(txt)}곳 | {total_dist_km:.1f}km | {fuel_cost:,}원")
        
        if v['orders']: folium.Marker(v['orders'][0]['coord'], icon=folium.Icon(color=c, icon='flag'), popup=v['name']).add_to(m)
        folium.PolyLine(path, color=c, weight=3, opacity=0.8).add_to(m)
        for o in v['orders']: folium.CircleMarker(o['coord'], radius=3, color=c, fill=True, popup=o['name']).add_to(m)

    m.save("최종_배차결과_지도.html")
    
    # [신규] 엑셀 저장 (시트 분리)
    with pd.ExcelWriter("최종_배차결과_통계포함.xlsx") as writer:
        df_detail.to_excel(writer, sheet_name="상세배차표", index=False)
        pd.DataFrame(summary_list).to_excel(writer, sheet_name="배차요약", index=False)
        
    print("\n🎉 모든 작업 완료! '배차요약' 시트에서 유류비를 확인하세요.")

except Exception as e:
    print(f"❌ 오류: {e}")