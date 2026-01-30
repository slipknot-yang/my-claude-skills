#!/usr/bin/env python3
"""
관세 이상 탐지 스크립트

Oracle DB에서 관세 신고 데이터의 이상 패턴을 탐지합니다:
- 과소신고 의심
- 단가 이상치
- HS코드 변경
- 종합 리스크 분석
- 고위험 업체 식별

Usage:
    python detect_anomalies.py
    python detect_anomalies.py --output anomaly_report.xlsx
    python detect_anomalies.py --year 2024 --threshold 1.5
"""

import oracledb
import pandas as pd
from datetime import datetime
import argparse
import sys

# DB 접속 정보
DB_CONFIG = {
    "user": "CLRIUSR",
    "password": "ntancisclri1!",
    "dsn": "211.239.120.42:3535/NTANCIS"
}

# HS 코드 매핑
HS4_NAMES = {
    '8518': '스피커/헤드폰/마이크',
    '8528': '모니터/TV',
    '8516': '전열기기',
    '9403': '가구',
    '8714': '차량부품(이륜차)',
    '8708': '차량부품(자동차)',
    '8703': '승용차',
    '6402': '고무/플라스틱 신발',
    '3926': '플라스틱 제품',
    '4202': '가방/지갑',
    '7323': '식탁/주방용품(철강)',
    '3924': '식탁/주방용품(플라스틱)',
    '8536': '전기회로 스위치',
    '3917': '플라스틱 관/호스',
    '8421': '원심분리기/필터',
}

# 국가 코드 매핑
COUNTRY_NAMES = {
    'CN': '중국',
    'JP': '일본',
    'IN': '인도',
    'AE': 'UAE',
    'PK': '파키스탄',
    'TH': '태국',
    'VN': '베트남',
    'MY': '말레이시아',
    'ID': '인도네시아',
    'TW': '대만',
}


def connect_db():
    """Oracle DB 연결"""
    print("🔗 DB 연결 중...")
    try:
        conn = oracledb.connect(**DB_CONFIG)
        print("✅ DB 연결 성공")
        return conn
    except oracledb.Error as e:
        print(f"❌ DB 연결 실패: {e}")
        sys.exit(1)


def detect_undervaluation(conn, year_filter="23", threshold=1.3):
    """과소신고 의심 건 탐지"""
    print(f"🔍 과소신고 탐지 중 (임계값: {threshold*100-100:.0f}% 이상)...")
    
    query = f"""
    SELECT 
        ASSD_HS_CD as HS_CODE,
        ORIG_CNTY_CD as COUNTRY,
        COUNT(*) as CNT,
        ROUND(AVG((ASSD_UT_USD_VAL - DCLD_UT_USD_VAL) / NULLIF(DCLD_UT_USD_VAL, 0) * 100), 1) as AVG_DIFF_PCT,
        SUM(ASSD_INVC_USD_AMT - DCLD_INVC_USD_AMT) as TOTAL_DIFF_USD,
        SUM(ASSD_INVC_USD_AMT) as TOTAL_VALUE_USD
    FROM CLRI_TANSAD_UT_PRC_M
    WHERE DEL_YN = 'N'
      AND DCLD_UT_USD_VAL > 0
      AND ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * {threshold}
      AND TANSAD_YY >= '{year_filter}'
    GROUP BY ASSD_HS_CD, ORIG_CNTY_CD
    HAVING COUNT(*) >= 10
    ORDER BY TOTAL_DIFF_USD DESC NULLS LAST
    FETCH FIRST 100 ROWS ONLY
    """
    
    df = pd.read_sql(query, conn)
    
    # 품목명, 국가명 추가
    df['HS4'] = df['HS_CODE'].str[:4]
    df['HS_NAME'] = df['HS4'].map(HS4_NAMES).fillna('기타')
    df['COUNTRY_NAME'] = df['COUNTRY'].map(COUNTRY_NAMES).fillna('기타')
    
    print(f"  → {len(df)}개 과소신고 의심 조합 탐지")
    return df


def detect_price_variance(conn, year_filter="23"):
    """단가 이상치 탐지"""
    print("📊 단가 이상치 탐지 중...")
    
    query = f"""
    SELECT 
        ASSD_HS_CD as HS_CODE,
        COUNT(*) as CNT,
        ROUND(AVG(ASSD_UT_USD_VAL), 2) as AVG_PRICE,
        ROUND(STDDEV(ASSD_UT_USD_VAL), 2) as STD_PRICE,
        ROUND(MIN(ASSD_UT_USD_VAL), 2) as MIN_PRICE,
        ROUND(MAX(ASSD_UT_USD_VAL), 2) as MAX_PRICE,
        ROUND(STDDEV(ASSD_UT_USD_VAL) / NULLIF(AVG(ASSD_UT_USD_VAL), 0) * 100, 1) as CV_PCT
    FROM CLRI_TANSAD_UT_PRC_M
    WHERE DEL_YN = 'N'
      AND ASSD_UT_USD_VAL > 0
      AND TANSAD_YY >= '{year_filter}'
    GROUP BY ASSD_HS_CD
    HAVING COUNT(*) >= 50 
       AND STDDEV(ASSD_UT_USD_VAL) > AVG(ASSD_UT_USD_VAL)
    ORDER BY STD_PRICE DESC
    FETCH FIRST 50 ROWS ONLY
    """
    
    df = pd.read_sql(query, conn)
    
    # 품목명 추가
    df['HS4'] = df['HS_CODE'].str[:4]
    df['HS_NAME'] = df['HS4'].map(HS4_NAMES).fillna('기타')
    
    print(f"  → {len(df)}개 단가 이상 품목 탐지")
    return df


def detect_hs_changes(conn, year_filter="23"):
    """HS코드 변경 탐지"""
    print("🔄 HS코드 변경 탐지 중...")
    
    query = f"""
    SELECT 
        DCLD_HS_CD as DECLARED_HS,
        ASSD_HS_CD as ASSESSED_HS,
        COUNT(*) as CNT,
        SUM(ASSD_INVC_USD_AMT) as TOTAL_VALUE_USD
    FROM CLRI_TANSAD_UT_PRC_M
    WHERE DEL_YN = 'N'
      AND DCLD_HS_CD IS NOT NULL
      AND ASSD_HS_CD IS NOT NULL
      AND DCLD_HS_CD != ASSD_HS_CD
      AND TANSAD_YY >= '{year_filter}'
    GROUP BY DCLD_HS_CD, ASSD_HS_CD
    HAVING COUNT(*) >= 20
    ORDER BY CNT DESC
    FETCH FIRST 50 ROWS ONLY
    """
    
    df = pd.read_sql(query, conn)
    
    # HS4 추출 및 명칭 추가
    df['DECLARED_HS4'] = df['DECLARED_HS'].str[:4]
    df['ASSESSED_HS4'] = df['ASSESSED_HS'].str[:4]
    
    print(f"  → {len(df)}개 HS코드 변경 패턴 탐지")
    return df


def calculate_risk_score(conn, year_filter="23"):
    """품목-국가 종합 리스크 분석"""
    print("⚠️ 종합 리스크 분석 중...")
    
    query = f"""
    WITH risk_data AS (
        SELECT 
            SUBSTR(ASSD_HS_CD, 1, 4) as HS4,
            ORIG_CNTY_CD,
            CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END as HS_CHANGED,
            CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.5 AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END as UNDERVALUED,
            ASSD_INVC_USD_AMT
        FROM CLRI_TANSAD_UT_PRC_M
        WHERE DEL_YN = 'N' AND TANSAD_YY >= '{year_filter}'
    )
    SELECT 
        HS4,
        ORIG_CNTY_CD as COUNTRY,
        COUNT(*) as TOTAL_CNT,
        SUM(HS_CHANGED) as HS_CHANGE_CNT,
        SUM(UNDERVALUED) as UNDERVALUE_CNT,
        ROUND(SUM(UNDERVALUED) * 100.0 / COUNT(*), 1) as UNDERVALUE_RATE,
        SUM(ASSD_INVC_USD_AMT) as TOTAL_VALUE_USD,
        ROUND(SUM(UNDERVALUED) * 3.0 / COUNT(*) * 100 + SUM(HS_CHANGED) * 2.0 / COUNT(*) * 100, 1) as RISK_SCORE
    FROM risk_data
    WHERE HS4 IS NOT NULL
    GROUP BY HS4, ORIG_CNTY_CD
    HAVING SUM(UNDERVALUED) >= 50 OR SUM(HS_CHANGED) >= 50
    ORDER BY RISK_SCORE DESC
    FETCH FIRST 50 ROWS ONLY
    """
    
    df = pd.read_sql(query, conn)
    
    # 명칭 추가
    df['HS_NAME'] = df['HS4'].map(HS4_NAMES).fillna('기타')
    df['COUNTRY_NAME'] = df['COUNTRY'].map(COUNTRY_NAMES).fillna('기타')
    
    # 리스크 등급
    df['RISK_GRADE'] = pd.cut(
        df['RISK_SCORE'], 
        bins=[-float('inf'), 30, 50, 80, float('inf')],
        labels=['NORMAL', 'LOW', 'MEDIUM', 'HIGH']
    )
    
    print(f"  → {len(df)}개 고위험 품목-국가 조합")
    return df


def identify_high_risk_importers(conn, year_filter="23"):
    """고위험 업체 식별"""
    print("🏢 고위험 업체 식별 중...")
    
    query = f"""
    SELECT 
        IMPPN_TIN as TIN,
        MAX(IMPPN_NM) as IMPORTER_NAME,
        COUNT(*) as TOTAL_CNT,
        SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.3 AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END) as UNDERVALUE_CNT,
        SUM(CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END) as HS_CHANGE_CNT,
        SUM(ASSD_INVC_USD_AMT) as TOTAL_VALUE_USD,
        ROUND(SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.3 AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as UNDERVALUE_RATE
    FROM CLRI_TANSAD_UT_PRC_M
    WHERE DEL_YN = 'N' AND TANSAD_YY >= '{year_filter}'
    GROUP BY IMPPN_TIN
    HAVING COUNT(*) >= 20
       AND (SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.3 AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END) >= 5
            OR SUM(CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END) >= 5)
    ORDER BY UNDERVALUE_CNT DESC
    FETCH FIRST 50 ROWS ONLY
    """
    
    df = pd.read_sql(query, conn)
    
    # 리스크 점수 계산
    df['RISK_SCORE'] = (df['UNDERVALUE_CNT'] * 3 + df['HS_CHANGE_CNT'] * 2) / df['TOTAL_CNT'] * 100
    df['RISK_SCORE'] = df['RISK_SCORE'].round(1)
    
    print(f"  → {len(df)}개 고위험 업체 식별")
    return df


def create_summary(df_underval, df_variance, df_hs_change, df_risk, df_importers):
    """요약 데이터 생성"""
    summary = {
        '지표': [
            '분석 기간',
            '과소신고 의심 조합 수',
            '과소신고 추정 총 차액 (USD)',
            '단가 이상 품목 수',
            'HS코드 변경 패턴 수',
            '고위험 품목-국가 조합',
            '고위험 업체 수',
            'HIGH 등급 조합 수',
            'MEDIUM 등급 조합 수',
        ],
        '값': [
            '2023-2024',
            f"{len(df_underval):,} 개",
            f"${df_underval['TOTAL_DIFF_USD'].sum():,.0f}" if 'TOTAL_DIFF_USD' in df_underval.columns else 'N/A',
            f"{len(df_variance):,} 개",
            f"{len(df_hs_change):,} 개",
            f"{len(df_risk):,} 개",
            f"{len(df_importers):,} 개",
            f"{(df_risk['RISK_GRADE'] == 'HIGH').sum():,} 개" if 'RISK_GRADE' in df_risk.columns else 'N/A',
            f"{(df_risk['RISK_GRADE'] == 'MEDIUM').sum():,} 개" if 'RISK_GRADE' in df_risk.columns else 'N/A',
        ]
    }
    return pd.DataFrame(summary)


def create_claude_prompts():
    """Claude for Excel 분석 프롬프트 생성"""
    prompts = {
        '시트명': [
            '과소신고_의심',
            '단가_이상',
            'HS코드_변경',
            '품목국가_리스크',
            '고위험_업체'
        ],
        'Claude 프롬프트': [
            '이 과소신고 데이터에서 패턴을 분석하고, 의도적 탈세와 단순 오류를 구분할 수 있는 기준을 제시해주세요.',
            '단가 편차가 큰 품목들의 특성을 분석하고, 가격 조작 가능성이 높은 품목을 식별해주세요.',
            'HS코드 변경 패턴을 분석하고, 관세 회피를 위한 의도적 분류 오류와 단순 실수를 구분해주세요.',
            '고위험 품목-국가 조합의 특성을 분석하고, 우선 점검 대상 TOP 10을 선정해주세요.',
            '업체별 리스크 패턴을 분석하고, 조사 우선순위와 예상 탈루 세액을 추정해주세요.'
        ],
        'Excel 함수': [
            '=CLAUDE("과소신고 패턴 분석", A1:H50)',
            '=CLAUDE("단가 이상 분석", A1:H50)',
            '=CLAUDE("HS변경 패턴 분석", A1:E50)',
            '=CLAUDE("리스크 조합 분석", A1:J50)',
            '=CLAUDE("고위험 업체 분석", A1:H50)'
        ]
    }
    return pd.DataFrame(prompts)


def save_to_excel(output_path, df_underval, df_variance, df_hs_change, df_risk, df_importers):
    """Excel 파일 저장"""
    print(f"\n📁 Excel 파일 생성 중: {output_path}")
    
    df_summary = create_summary(df_underval, df_variance, df_hs_change, df_risk, df_importers)
    df_prompts = create_claude_prompts()
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='요약', index=False)
        df_underval.to_excel(writer, sheet_name='과소신고_의심', index=False)
        df_variance.to_excel(writer, sheet_name='단가_이상', index=False)
        df_hs_change.to_excel(writer, sheet_name='HS코드_변경', index=False)
        df_risk.to_excel(writer, sheet_name='품목국가_리스크', index=False)
        df_importers.to_excel(writer, sheet_name='고위험_업체', index=False)
        df_prompts.to_excel(writer, sheet_name='Claude_분석_가이드', index=False)
    
    print(f"✅ Excel 파일 저장 완료")


def print_alert_summary(df_underval, df_risk, df_importers):
    """콘솔에 알림 출력"""
    print("\n" + "="*60)
    print("🚨 이상 탐지 알림 요약")
    print("="*60)
    
    # 과소신고 TOP 5
    print("\n📌 과소신고 의심 TOP 5 (품목-국가):")
    for i, row in df_underval.head(5).iterrows():
        print(f"  {i+1}. {row['HS4']} ({row['HS_NAME']}) + {row['COUNTRY']} ({row['COUNTRY_NAME']})")
        print(f"      → {row['CNT']:,}건, 평균 차이 {row['AVG_DIFF_PCT']:.1f}%, 총 차액 ${row['TOTAL_DIFF_USD']:,.0f}")
    
    # 고위험 조합 TOP 5
    print("\n🔴 고위험 품목-국가 TOP 5:")
    if 'RISK_SCORE' in df_risk.columns:
        for i, row in df_risk.head(5).iterrows():
            print(f"  {i+1}. {row['HS4']} ({row['HS_NAME']}) + {row['COUNTRY']} ({row['COUNTRY_NAME']})")
            print(f"      → 리스크 점수: {row['RISK_SCORE']:.1f}, 과소신고율: {row['UNDERVALUE_RATE']:.1f}%")
    
    # 고위험 업체 TOP 5
    print("\n🏢 고위험 업체 TOP 5:")
    for i, row in df_importers.head(5).iterrows():
        name = row['IMPORTER_NAME'][:20] if pd.notna(row['IMPORTER_NAME']) else 'N/A'
        print(f"  {i+1}. {row['TIN']} ({name}...)")
        print(f"      → 과소신고 {row['UNDERVALUE_CNT']:,}건 ({row['UNDERVALUE_RATE']:.1f}%), 총 ${row['TOTAL_VALUE_USD']:,.0f}")
    
    print("\n" + "="*60)


def main():
    parser = argparse.ArgumentParser(description='관세 이상 탐지')
    parser.add_argument('--output', '-o', default='customs_anomaly_report.xlsx',
                        help='출력 Excel 파일 경로')
    parser.add_argument('--year', '-y', default='23',
                        help='분석 시작 연도 (2자리, 예: 23)')
    parser.add_argument('--threshold', '-t', type=float, default=1.3,
                        help='과소신고 임계값 (기본: 1.3 = 30%)')
    args = parser.parse_args()
    
    print("🚀 관세 이상 탐지 시작")
    print(f"⏰ 시작 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📅 분석 기간: 20{args.year}년 이후")
    print(f"📊 과소신고 임계값: {(args.threshold-1)*100:.0f}% 이상")
    
    # DB 연결
    conn = connect_db()
    
    try:
        # 이상 탐지 실행
        df_underval = detect_undervaluation(conn, args.year, args.threshold)
        df_variance = detect_price_variance(conn, args.year)
        df_hs_change = detect_hs_changes(conn, args.year)
        df_risk = calculate_risk_score(conn, args.year)
        df_importers = identify_high_risk_importers(conn, args.year)
        
        # Excel 저장
        save_to_excel(args.output, df_underval, df_variance, df_hs_change, df_risk, df_importers)
        
        # 알림 출력
        print_alert_summary(df_underval, df_risk, df_importers)
        
        print(f"\n✅ 이상 탐지 완료!")
        print(f"📁 결과 파일: {args.output}")
        print(f"\n💡 Claude for Excel에서 'Claude_분석_가이드' 시트의 프롬프트를 활용하세요.")
        
    finally:
        conn.close()
        print("🔌 DB 연결 종료")


if __name__ == "__main__":
    main()
