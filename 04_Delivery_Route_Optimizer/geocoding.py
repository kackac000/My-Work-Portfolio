import pandas as pd
import requests
import time

# ==========================================
# [설정] 카카오 API 키 (여기에 키를 다시 넣어주세요)
KAKAO_API_KEY = "d42be7397b700f464a6a73065cb0182c" 
# ==========================================

# 매핑 설정 (SAP 헤더 -> 프로그램 변수)
COLUMN_MAPPING = {
    '거래처명': 'name', 
    '주소': 'address', 
    '박스수량': 'box_qty', 
    '봉지수량': 'bag_qty',
    '하차시간': 'service_time', 
    '오픈시간': 'time_window_start', 
    '마감시간': 'time_window_end', 
    '비고': 'memo'
}

def get_lat_lon(address, index, total):
    """
    주소 -> 좌표 변환 함수
    특징: 변환할 때마다 몇 번째인지 화면에 바로 출력함
    """
    if pd.isna(address) or str(address).strip() == "":
        print(f"[{index}/{total}] ❌ 주소없음")
        return "주소없음"
    
    url = 'https://dapi.kakao.com/v2/local/search/address.json'
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    params = {'query': str(address)}

    try:
        response = requests.get(url, headers=headers, params=params, timeout=5)
        if response.status_code == 200:
            result = response.json()
            if result['documents']:
                x = result['documents'][0]['x']
                y = result['documents'][0]['y']
                # [핵심] 성공하면 바로 출력!
                print(f"[{index}/{total}] ✅ 성공: {str(address)[:7]}... -> {x},{y}")
                return f"{x},{y}"
            else:
                print(f"[{index}/{total}] ⚠️ 실패: {address} (검색안됨)")
                return "검색실패"
        else:
            print(f"[{index}/{total}] 🚫 에러: {response.status_code}")
            return "API에러"
    except Exception as e:
        print(f"[{index}/{total}] 💥 오류: {e}")
        return f"오류:{e}"

# --- 메인 실행 부분 ---
print("🚀 좌표 변환기 가동! (실시간 중계 모드)")

try:
    # 1. 파일 읽기
    print("📂 엑셀 파일을 읽고 있습니다...")
    df = pd.read_excel("배차데이터_양식.xlsx", sheet_name="주문목록")
    
    # 헤더 이름 바꾸기
    df.rename(columns=COLUMN_MAPPING, inplace=True)
    
    total_count = len(df)
    print(f"✅ 총 {total_count}건의 데이터를 발견했습니다. 변환을 시작합니다!\n")

    # 2. 한 줄씩 순서대로 변환 (여기가 바뀐 부분입니다)
    results = []
    for i, row in df.iterrows():
        # 현재 순서(i+1)와 전체 개수(total_count)를 같이 넘겨줍니다.
        current_addr = row.get('address')
        res = get_lat_lon(current_addr, i+1, total_count)
        results.append(res)
    
    # 결과를 데이터프레임에 넣기
    df['좌표'] = results

    # 3. 물량 계산 및 저장
    df['box_qty'] = df['box_qty'].fillna(0)
    df['bag_qty'] = df['bag_qty'].fillna(0)
    df['total_load'] = (df['box_qty'] * 1.0) + (df['bag_qty'] * 0.2)

    output_filename = "배차데이터_좌표완료.xlsx"
    with pd.ExcelWriter(output_filename) as writer:
        df.to_excel(writer, sheet_name="주문목록", index=False)
        print(f"\n🎉 고생하셨습니다! '{output_filename}' 파일 저장 완료!")

except FileNotFoundError:
    print("\n❌ 엑셀 파일이 없습니다. 파일명을 확인해주세요.")
except Exception as e:
    print(f"\n❌ 오류 발생: {e}")
