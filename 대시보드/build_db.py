import duckdb
import os
import shutil
from datetime import datetime

# 경로 설정
BASE_DIR = r'C:\My_Project'
INPUT_DIR = os.path.join(BASE_DIR, 'Input')
ARCHIVE_DIR = os.path.join(BASE_DIR, 'Archive')
MASTER_DIR = os.path.join(BASE_DIR, 'Master')
DB_PATH = os.path.join(BASE_DIR, 'pharmacy_data.db')

def get_columns(con, view_name):
    return [row[1] for row in con.execute(f"PRAGMA table_info('{view_name}')").fetchall()]

def process_excel_to_db():
    con = duckdb.connect(DB_PATH)
    con.execute("INSTALL spatial; LOAD spatial;")
    
    # 1. 초기 테이블 세팅 (컬럼 16개로 확장: Storage_Loc 추가)
    con.execute("CREATE TABLE IF NOT EXISTS file_log (filename TEXT PRIMARY KEY, processed_at TIMESTAMP)")
    con.execute("CREATE SEQUENCE IF NOT EXISTS seq_id START 1")
    con.execute("""
        CREATE TABLE IF NOT EXISTS Movement_Master (
            ID INTEGER, Date DATE, Time TIME, Ref_No TEXT, 
            Move_Type TEXT, 
            Move_Desc TEXT,         -- [로직] 601+1006='약가인하' 반영 예정
            Storage_Loc TEXT,       -- [추가] 보관장소 컬럼
            Plant_Code TEXT, Partner_Code TEXT, 
            Partner_Name TEXT,      -- [로직] 미등록 거래처 이름 부여 예정
            Partner_Address TEXT, 
            Product_Code TEXT, Product_Name TEXT, Batch TEXT, Qty DOUBLE, Amount DOUBLE
        )
    """)

    con.execute(f"CREATE OR REPLACE VIEW v_logic AS SELECT * FROM st_read('{MASTER_DIR}/Logic_Movement.xlsx')")
    con.execute(f"CREATE OR REPLACE VIEW v_partner AS SELECT * FROM st_read('{MASTER_DIR}/Master_Partner.xlsx')")
    
    partner_cols = get_columns(con, 'v_partner')
    logic_cols = get_columns(con, 'v_logic')

    L_TYPE = "Movement Type" if "Movement Type" in logic_cols else logic_cols[0]
    L_DESC = "Movement Type Description" if "Movement Type Description" in logic_cols else logic_cols[2]
    P_CUST = "Customer" if "Customer" in partner_cols else partner_cols[0]
    P_NAME = "Name 1" if "Name 1" in partner_cols else partner_cols[1]
    P_ADDR = "Street" if "Street" in partner_cols else (partner_cols[2] if len(partner_cols) > 2 else "Street")

    files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(('.xlsx', '.xls'))]
    
    for file in files:
        if con.execute("SELECT 1 FROM file_log WHERE filename = ?", [file]).fetchone():
            shutil.move(os.path.join(INPUT_DIR, file), os.path.join(ARCHIVE_DIR, file))
            continue

        print(f"📦 {file} 정제 중...")
        con.execute(f"CREATE OR REPLACE VIEW v_raw AS SELECT * FROM st_read('{INPUT_DIR}/{file}')")
        
        # 4. 정제 SQL (핵심 로직 2가지 반영)
        sql_query = f"""
        INSERT INTO Movement_Master
        SELECT 
            nextval('seq_id'),
            "Posting Date"::DATE,
            TIME '00:00:00' + (CAST("Time of Entry" AS DOUBLE) * INTERVAL '1 day'),
            COALESCE(CAST(Reference AS TEXT), CAST("Purchase order" AS TEXT)),
            CAST("Movement Type" AS TEXT),
            -- [로직 1] 601 타입 중 보관장소 1006은 '약가인하'로 별도 분류
            CASE 
                WHEN CAST(raw."Movement Type" AS TEXT) = '601' AND CAST(raw."Storage Location" AS TEXT) = '1006' THEN '약가인하'
                ELSE logic."{L_DESC}"
            END,
            CAST(raw."Storage Location" AS TEXT), -- Storage_Loc 저장
            CASE 
                WHEN CAST(Plant AS TEXT) IN ('3115', '3125') THEN '대구'
                WHEN CAST(Plant AS TEXT) IN ('3116', '3126') THEN '부산'
                WHEN CAST(Plant AS TEXT) IN ('3117', '3127') THEN '대전'
                WHEN CAST(Plant AS TEXT) IN ('3118', '3128', '3119', '3129') THEN '화성'
                ELSE partner_p."{P_NAME}"
            END,
            COALESCE(CAST(raw.Customer AS TEXT), CAST(raw.Supplier AS TEXT)),
            -- [로직 2] 거래처명이 없을 경우 '미등록(코드)'으로 표시하여 공란 뭉침 방지
            COALESCE(partner."{P_NAME}", '미등록(' || COALESCE(CAST(raw.Customer AS TEXT), CAST(raw.Supplier AS TEXT), 'N/A') || ')'),
            partner."{P_ADDR}",
            CAST(Material AS TEXT),
            "Material Description",
            Batch,
            ABS(CAST(Quantity AS DOUBLE)),
            ABS(CAST("Amt.in loc.cur." AS DOUBLE))
        FROM v_raw AS raw
        LEFT JOIN v_logic AS logic ON CAST(raw."Movement Type" AS TEXT) = CAST(logic."{L_TYPE}" AS TEXT)
        LEFT JOIN v_partner AS partner ON COALESCE(CAST(raw.Customer AS TEXT), CAST(raw.Supplier AS TEXT)) = CAST(partner."{P_CUST}" AS TEXT)
        LEFT JOIN v_partner AS partner_p ON CAST(raw.Plant AS TEXT) = CAST(partner_p."{P_CUST}" AS TEXT)
        """
        
        try:
            con.execute(sql_query)
            con.execute("INSERT INTO file_log VALUES (?, ?)", [file, datetime.now()])
            shutil.move(os.path.join(INPUT_DIR, file), os.path.join(ARCHIVE_DIR, file))
            print(f"✅ {file} 처리 완료")
            con.execute("CREATE INDEX IF NOT EXISTS idx_batch ON Movement_Master (Batch)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_ref ON Movement_Master (Ref_No)")
        except Exception as e:
            con.close()
            raise e
    con.close()

if __name__ == "__main__":
    process_excel_to_db()
