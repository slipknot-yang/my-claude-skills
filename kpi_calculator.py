#!/usr/bin/env python3
"""
WCO PMM 기반 관세 KPI 계산 모듈

WCO Performance Measurement Model (PMM) 4차원:
1. Trade Facilitation (무역 원활화)
2. Revenue Collection (세수 확보)
3. Risk Management & Enforcement (위험관리 및 집행)
4. Organizational Development (조직 발전)

References:
- WCO Customs Risk Management Compendium (2022)
- KCS 관세연감 지표 체계
- UN Comtrade 분석 프레임워크
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass
from enum import Enum


class KPICategory(Enum):
    """WCO PMM 4차원 분류"""
    TRADE_FACILITATION = "Trade Facilitation"
    REVENUE_COLLECTION = "Revenue Collection"
    RISK_MANAGEMENT = "Risk Management"
    ORGANIZATIONAL = "Organizational"


@dataclass
class KPIDefinition:
    """KPI 정의"""
    code: str
    name_en: str
    name_kr: str
    category: KPICategory
    unit: str
    description: str
    formula: str
    benchmark: Optional[float] = None
    target: Optional[float] = None
    direction: str = "higher"  # higher=높을수록 좋음, lower=낮을수록 좋음


# WCO PMM 기반 KPI 정의 (16개)
KPI_DEFINITIONS: Dict[str, KPIDefinition] = {
    # === Trade Facilitation (4개) ===
    "TF001": KPIDefinition(
        code="TF001",
        name_en="Average Clearance Time",
        name_kr="평균 통관시간",
        category=KPICategory.TRADE_FACILITATION,
        unit="hours",
        description="신고부터 수리까지 평균 소요 시간",
        formula="AVG(수리일시 - 신고일시)",
        benchmark=24.0,
        target=12.0,
        direction="lower"
    ),
    "TF002": KPIDefinition(
        code="TF002",
        name_en="Pre-arrival Processing Rate",
        name_kr="사전신고처리율",
        category=KPICategory.TRADE_FACILITATION,
        unit="%",
        description="입항 전 사전신고 건수 비율",
        formula="사전신고건수 / 총신고건수 × 100",
        benchmark=30.0,
        target=50.0,
        direction="higher"
    ),
    "TF003": KPIDefinition(
        code="TF003",
        name_en="Green Lane Rate",
        name_kr="녹색통로비율",
        category=KPICategory.TRADE_FACILITATION,
        unit="%",
        description="무검사 통관 비율",
        formula="무검사건수 / 총신고건수 × 100",
        benchmark=70.0,
        target=85.0,
        direction="higher"
    ),
    "TF004": KPIDefinition(
        code="TF004",
        name_en="Electronic Declaration Rate",
        name_kr="전자신고율",
        category=KPICategory.TRADE_FACILITATION,
        unit="%",
        description="전자신고 처리 비율",
        formula="전자신고건수 / 총신고건수 × 100",
        benchmark=95.0,
        target=99.0,
        direction="higher"
    ),
    
    # === Revenue Collection (4개) ===
    "RC001": KPIDefinition(
        code="RC001",
        name_en="Collection Efficiency Ratio",
        name_kr="징수효율",
        category=KPICategory.REVENUE_COLLECTION,
        unit="%",
        description="확정세액 대비 실제 징수액 비율",
        formula="실제징수액 / 확정세액 × 100",
        benchmark=95.0,
        target=99.0,
        direction="higher"
    ),
    "RC002": KPIDefinition(
        code="RC002",
        name_en="Duty Assessment Accuracy",
        name_kr="세액심사정확도",
        category=KPICategory.REVENUE_COLLECTION,
        unit="%",
        description="심사 후 세액 변경이 없는 비율",
        formula="(1 - 세액변경건수/심사건수) × 100",
        benchmark=90.0,
        target=95.0,
        direction="higher"
    ),
    "RC003": KPIDefinition(
        code="RC003",
        name_en="Post-clearance Audit Coverage",
        name_kr="사후심사커버리지",
        category=KPICategory.REVENUE_COLLECTION,
        unit="%",
        description="사후심사 대상 업체 비율",
        formula="심사업체수 / 총수입업체수 × 100",
        benchmark=5.0,
        target=10.0,
        direction="higher"
    ),
    "RC004": KPIDefinition(
        code="RC004",
        name_en="Revenue Growth Rate",
        name_kr="세수증가율",
        category=KPICategory.REVENUE_COLLECTION,
        unit="%",
        description="전년 대비 관세 수입 증가율",
        formula="(당기세수 - 전기세수) / 전기세수 × 100",
        benchmark=3.0,
        target=5.0,
        direction="higher"
    ),
    
    # === Risk Management (5개) ===
    "RM001": KPIDefinition(
        code="RM001",
        name_en="Selectivity Rate",
        name_kr="선별율",
        category=KPICategory.RISK_MANAGEMENT,
        unit="%",
        description="검사 대상 선별 비율",
        formula="검사선별건수 / 총신고건수 × 100",
        benchmark=15.0,
        target=10.0,
        direction="lower"
    ),
    "RM002": KPIDefinition(
        code="RM002",
        name_en="Hit Rate",
        name_kr="적발율",
        category=KPICategory.RISK_MANAGEMENT,
        unit="%",
        description="검사 건 중 위반 적발 비율",
        formula="위반적발건수 / 검사건수 × 100",
        benchmark=20.0,
        target=30.0,
        direction="higher"
    ),
    "RM003": KPIDefinition(
        code="RM003",
        name_en="Compliance Rate",
        name_kr="신고정확도",
        category=KPICategory.RISK_MANAGEMENT,
        unit="%",
        description="정확한 신고 비율",
        formula="정확신고건수 / 총신고건수 × 100",
        benchmark=85.0,
        target=95.0,
        direction="higher"
    ),
    "RM004": KPIDefinition(
        code="RM004",
        name_en="Undervaluation Detection Rate",
        name_kr="과소신고탐지율",
        category=KPICategory.RISK_MANAGEMENT,
        unit="%",
        description="과소신고 의심 건 탐지 비율",
        formula="과소신고적발건수 / 과소신고의심건수 × 100",
        benchmark=30.0,
        target=50.0,
        direction="higher"
    ),
    "RM005": KPIDefinition(
        code="RM005",
        name_en="HS Misclassification Rate",
        name_kr="품목분류오류율",
        category=KPICategory.RISK_MANAGEMENT,
        unit="%",
        description="HS코드 변경 비율",
        formula="품목변경건수 / 총신고건수 × 100",
        benchmark=5.0,
        target=2.0,
        direction="lower"
    ),
    
    # === Organizational (3개) ===
    "OD001": KPIDefinition(
        code="OD001",
        name_en="HHI Concentration Index",
        name_kr="HHI집중도",
        category=KPICategory.ORGANIZATIONAL,
        unit="index",
        description="수입 품목/국가 집중도 (Herfindahl-Hirschman Index)",
        formula="Σ(시장점유율²)",
        benchmark=1500,
        target=1000,
        direction="lower"
    ),
    "OD002": KPIDefinition(
        code="OD002",
        name_en="MoM Volatility",
        name_kr="월간변동성",
        category=KPICategory.ORGANIZATIONAL,
        unit="%",
        description="월별 세수 변동 표준편차",
        formula="STDDEV(월별세수) / AVG(월별세수) × 100",
        benchmark=15.0,
        target=10.0,
        direction="lower"
    ),
    "OD003": KPIDefinition(
        code="OD003",
        name_en="YoY Growth Consistency",
        name_kr="연간성장일관성",
        category=KPICategory.ORGANIZATIONAL,
        unit="score",
        description="연간 성장률 일관성 (1-5점)",
        formula="성장률 연속성 기반 점수",
        benchmark=3.0,
        target=4.0,
        direction="higher"
    ),
}


class KPICalculator:
    """WCO PMM 기반 KPI 계산기"""
    
    def __init__(self, conn):
        """
        Args:
            conn: Oracle DB 연결 객체
        """
        self.conn = conn
        self.cache: Dict[str, Any] = {}
    
    def get_definition(self, kpi_code: str) -> Optional[KPIDefinition]:
        """KPI 정의 조회"""
        return KPI_DEFINITIONS.get(kpi_code)
    
    def get_all_definitions(self) -> Dict[str, KPIDefinition]:
        """모든 KPI 정의 조회"""
        return KPI_DEFINITIONS
    
    def get_definitions_by_category(self, category: KPICategory) -> Dict[str, KPIDefinition]:
        """카테고리별 KPI 정의 조회"""
        return {k: v for k, v in KPI_DEFINITIONS.items() if v.category == category}
    
    # === Revenue Collection KPIs ===
    
    def calc_revenue_by_period(self, period: str = 'yearly') -> pd.DataFrame:
        """기간별 세수 현황
        
        Args:
            period: 'yearly', 'monthly', 'quarterly'
        """
        if period == 'yearly':
            sql = """
                SELECT 
                    '20' || TANSAD_YY as period,
                    COUNT(*) as declaration_count,
                    SUM(ITM_TAX_AMT) as total_tax,
                    SUM(ITM_INVC_USD_AMT) as total_value_usd,
                    AVG(ITM_TAX_AMT) as avg_tax_per_item,
                    COUNT(DISTINCT SUBSTR(HS_CD, 1, 2)) as hs_chapter_count,
                    COUNT(DISTINCT ORIG_CNTY_CD) as country_count
                FROM CLRI_TANSAD_ITM_D
                WHERE DEL_YN = 'N' AND TANSAD_YY >= '20'
                GROUP BY TANSAD_YY
                ORDER BY TANSAD_YY DESC
            """
        elif period == 'monthly':
            sql = """
                SELECT 
                    TO_CHAR(FRST_RGSR_DTM, 'YYYY-MM') as period,
                    COUNT(*) as declaration_count,
                    SUM(ITM_TAX_AMT) as total_tax,
                    SUM(ITM_INVC_USD_AMT) as total_value_usd,
                    AVG(ITM_TAX_AMT) as avg_tax_per_item
                FROM CLRI_TANSAD_ITM_D
                WHERE DEL_YN = 'N' 
                  AND FRST_RGSR_DTM >= ADD_MONTHS(SYSDATE, -36)
                GROUP BY TO_CHAR(FRST_RGSR_DTM, 'YYYY-MM')
                ORDER BY period DESC
            """
        elif period == 'quarterly':
            sql = """
                SELECT 
                    '20' || TANSAD_YY || '-Q' || TO_CHAR(FRST_RGSR_DTM, 'Q') as period,
                    COUNT(*) as declaration_count,
                    SUM(ITM_TAX_AMT) as total_tax,
                    SUM(ITM_INVC_USD_AMT) as total_value_usd
                FROM CLRI_TANSAD_ITM_D
                WHERE DEL_YN = 'N' AND TANSAD_YY >= '22'
                GROUP BY TANSAD_YY, TO_CHAR(FRST_RGSR_DTM, 'Q')
                ORDER BY period DESC
            """
        else:
            raise ValueError(f"Unknown period: {period}")
        
        df = pd.read_sql(sql, self.conn)
        # Oracle은 대문자 컬럼명 반환 -> 소문자로 변환
        df.columns = df.columns.str.lower()
        return df
    
    def calc_yoy_growth(self) -> pd.DataFrame:
        """연간 성장률 (YoY Growth Rate - RC004)"""
        df = self.calc_revenue_by_period('yearly')
        df = df.sort_values('period')
        df['prev_tax'] = df['total_tax'].shift(1)
        df['yoy_growth_pct'] = ((df['total_tax'] - df['prev_tax']) / df['prev_tax'] * 100).round(1)
        df['prev_count'] = df['declaration_count'].shift(1)
        df['yoy_count_growth_pct'] = ((df['declaration_count'] - df['prev_count']) / df['prev_count'] * 100).round(1)
        return df.dropna()
    
    def calc_mom_growth(self) -> pd.DataFrame:
        """월간 성장률 (MoM Growth)"""
        df = self.calc_revenue_by_period('monthly')
        df = df.sort_values('period')
        df['prev_tax'] = df['total_tax'].shift(1)
        df['mom_growth_pct'] = ((df['total_tax'] - df['prev_tax']) / df['prev_tax'] * 100).round(1)
        return df.dropna()
    
    def calc_volatility(self) -> Dict[str, float]:
        """월간 변동성 (OD002)"""
        df = self.calc_revenue_by_period('monthly')
        mean_tax = df['total_tax'].mean()
        std_tax = df['total_tax'].std()
        cv = (std_tax / mean_tax * 100) if mean_tax > 0 else 0
        return {
            'mean_monthly_tax': mean_tax,
            'std_monthly_tax': std_tax,
            'coefficient_of_variation': round(cv, 2),
            'volatility_rating': 'Low' if cv < 10 else 'Medium' if cv < 20 else 'High'
        }
    
    # === Risk Management KPIs ===
    
    def calc_undervaluation_stats(self, threshold: float = 1.3) -> pd.DataFrame:
        """과소신고 통계 (RM004)
        
        Args:
            threshold: 심사가격/신고가격 비율 임계값 (1.3 = 30% 이상 차이)
        """
        sql = f"""
            SELECT 
                '20' || TANSAD_YY as period,
                COUNT(*) as total_count,
                SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * {threshold} 
                         AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END) as underval_count,
                ROUND(SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * {threshold} 
                               AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2) as underval_rate,
                SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * {threshold} 
                         AND DCLD_UT_USD_VAL > 0 
                         THEN ASSD_INVC_USD_AMT - DCLD_INVC_USD_AMT ELSE 0 END) as estimated_loss_usd
            FROM CLRI_TANSAD_UT_PRC_M
            WHERE DEL_YN = 'N' AND TANSAD_YY >= '20'
            GROUP BY TANSAD_YY
            ORDER BY period DESC
        """
        df = pd.read_sql(sql, self.conn)
        df.columns = df.columns.str.lower()
        return df
    
    def calc_hs_misclassification_rate(self) -> pd.DataFrame:
        """품목분류 오류율 (RM005)"""
        sql = """
            SELECT 
                '20' || TANSAD_YY as period,
                COUNT(*) as total_count,
                SUM(CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END) as misclass_count,
                ROUND(SUM(CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2) as misclass_rate
            FROM CLRI_TANSAD_UT_PRC_M
            WHERE DEL_YN = 'N' AND TANSAD_YY >= '20'
            GROUP BY TANSAD_YY
            ORDER BY period DESC
        """
        df = pd.read_sql(sql, self.conn)
        df.columns = df.columns.str.lower()
        return df
    
    def calc_risk_score_by_hs_country(self, min_count: int = 50) -> pd.DataFrame:
        """품목-국가별 리스크 점수"""
        sql = f"""
            WITH base AS (
                SELECT 
                    SUBSTR(ASSD_HS_CD, 1, 4) as hs4,
                    ORIG_CNTY_CD as country,
                    COUNT(*) as total_count,
                    SUM(CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END) as hs_change_count,
                    SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.3 
                             AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END) as underval_count,
                    SUM(ASSD_INVC_USD_AMT) as total_value
                FROM CLRI_TANSAD_UT_PRC_M
                WHERE DEL_YN = 'N' AND TANSAD_YY >= '23'
                GROUP BY SUBSTR(ASSD_HS_CD, 1, 4), ORIG_CNTY_CD
                HAVING COUNT(*) >= {min_count}
            )
            SELECT 
                hs4,
                country,
                total_count,
                underval_count,
                ROUND(underval_count * 100.0 / total_count, 1) as underval_rate,
                hs_change_count,
                ROUND(hs_change_count * 100.0 / total_count, 1) as hs_change_rate,
                total_value,
                -- Risk Score = Underval Weight(3) + HS Change Weight(2) + Volume Weight(1)
                ROUND(
                    (underval_count * 3.0 / total_count * 100) + 
                    (hs_change_count * 2.0 / total_count * 100) +
                    (CASE WHEN total_value > 100000000 THEN 10 
                          WHEN total_value > 10000000 THEN 5 
                          ELSE 1 END)
                , 1) as risk_score
            FROM base
            WHERE underval_count > 0 OR hs_change_count > 0
            ORDER BY risk_score DESC
            FETCH FIRST 100 ROWS ONLY
        """
        df = pd.read_sql(sql, self.conn)
        df.columns = df.columns.str.lower()
        return df
    
    # === Organizational KPIs ===
    
    def calc_hhi_by_dimension(self, dimension: str = 'hs2') -> Dict[str, Any]:
        """HHI 집중도 (OD001)
        
        Args:
            dimension: 'hs2' (품목), 'country' (국가)
        """
        if dimension == 'hs2':
            sql = """
                SELECT 
                    SUBSTR(HS_CD, 1, 2) as category,
                    SUM(ITM_TAX_AMT) as value
                FROM CLRI_TANSAD_ITM_D
                WHERE DEL_YN = 'N' AND TANSAD_YY >= '23'
                GROUP BY SUBSTR(HS_CD, 1, 2)
            """
        elif dimension == 'country':
            sql = """
                SELECT 
                    ORIG_CNTY_CD as category,
                    SUM(ITM_INVC_USD_AMT) as value
                FROM CLRI_TANSAD_ITM_D
                WHERE DEL_YN = 'N' AND TANSAD_YY >= '23' AND ORIG_CNTY_CD IS NOT NULL
                GROUP BY ORIG_CNTY_CD
            """
        else:
            raise ValueError(f"Unknown dimension: {dimension}")
        
        df = pd.read_sql(sql, self.conn)
        df.columns = df.columns.str.lower()
        total = df['value'].sum()
        
        if total > 0:
            df['share'] = df['value'] / total
            df['share_squared'] = df['share'] ** 2
            hhi = df['share_squared'].sum() * 10000  # HHI는 10,000 스케일
        else:
            hhi = 0
        
        # HHI 해석
        if hhi < 1000:
            concentration = 'Low (Competitive)'
        elif hhi < 1800:
            concentration = 'Moderate'
        else:
            concentration = 'High (Concentrated)'
        
        return {
            'dimension': dimension,
            'hhi': round(hhi, 0),
            'concentration_level': concentration,
            'top_5_share': round(df.nlargest(5, 'value')['share'].sum() * 100, 1),
            'total_categories': len(df)
        }
    
    def calc_pareto_analysis(self, dimension: str = 'hs2', value_col: str = 'tax') -> pd.DataFrame:
        """파레토 분석 (80/20 규칙)"""
        if dimension == 'hs2':
            if value_col == 'tax':
                sql = """
                    SELECT 
                        SUBSTR(HS_CD, 1, 2) as category,
                        SUM(ITM_TAX_AMT) as value
                    FROM CLRI_TANSAD_ITM_D
                    WHERE DEL_YN = 'N' AND TANSAD_YY >= '23'
                    GROUP BY SUBSTR(HS_CD, 1, 2)
                    ORDER BY value DESC
                """
            else:
                sql = """
                    SELECT 
                        SUBSTR(HS_CD, 1, 2) as category,
                        SUM(ITM_INVC_USD_AMT) as value
                    FROM CLRI_TANSAD_ITM_D
                    WHERE DEL_YN = 'N' AND TANSAD_YY >= '23'
                    GROUP BY SUBSTR(HS_CD, 1, 2)
                    ORDER BY value DESC
                """
        elif dimension == 'country':
            sql = """
                SELECT 
                    ORIG_CNTY_CD as category,
                    SUM(ITM_INVC_USD_AMT) as value
                FROM CLRI_TANSAD_ITM_D
                WHERE DEL_YN = 'N' AND TANSAD_YY >= '23' AND ORIG_CNTY_CD IS NOT NULL
                GROUP BY ORIG_CNTY_CD
                ORDER BY value DESC
            """
        else:
            raise ValueError(f"Unknown dimension: {dimension}")
        
        df = pd.read_sql(sql, self.conn)
        df.columns = df.columns.str.lower()
        total = df['value'].sum()
        df['share_pct'] = (df['value'] / total * 100).round(2)
        df['cumulative_pct'] = df['share_pct'].cumsum().round(2)
        df['rank'] = range(1, len(df) + 1)
        df['pareto_zone'] = df['cumulative_pct'].apply(
            lambda x: 'A (Top 80%)' if x <= 80 else 'B (80-95%)' if x <= 95 else 'C (Bottom 5%)'
        )
        
        return df
    
    # === 통합 대시보드 KPIs ===
    
    def calc_executive_summary(self) -> Dict[str, Any]:
        """경영진 대시보드 요약 KPIs"""
        # 기본 통계
        basic_sql = """
            SELECT 
                COUNT(*) as total_declarations,
                SUM(ITM_TAX_AMT) as total_tax,
                SUM(ITM_INVC_USD_AMT) as total_value_usd,
                COUNT(DISTINCT SUBSTR(HS_CD, 1, 2)) as hs_chapters,
                COUNT(DISTINCT ORIG_CNTY_CD) as countries
            FROM CLRI_TANSAD_ITM_D
            WHERE DEL_YN = 'N' AND TANSAD_YY >= '23'
        """
        basic_df = pd.read_sql(basic_sql, self.conn)
        basic_df.columns = basic_df.columns.str.lower()
        basic = basic_df.iloc[0]
        
        # YoY 성장률
        yoy = self.calc_yoy_growth()
        latest_yoy = yoy.iloc[-1]['yoy_growth_pct'] if len(yoy) > 0 else 0
        
        # 변동성
        vol = self.calc_volatility()
        
        # HHI
        hhi_hs = self.calc_hhi_by_dimension('hs2')
        hhi_country = self.calc_hhi_by_dimension('country')
        
        # 과소신고
        underval = self.calc_undervaluation_stats()
        latest_underval = underval.iloc[0]['underval_rate'] if len(underval) > 0 else 0
        
        return {
            'period': '2023-2026',
            'total_declarations': int(basic['total_declarations']),
            'total_tax_krw': float(basic['total_tax']),
            'total_value_usd': float(basic['total_value_usd']),
            'hs_chapters': int(basic['hs_chapters']),
            'countries': int(basic['countries']),
            'yoy_growth_pct': float(latest_yoy),
            'volatility_cv': vol['coefficient_of_variation'],
            'hhi_commodity': hhi_hs['hhi'],
            'hhi_country': hhi_country['hhi'],
            'underval_rate': float(latest_underval),
            'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
    
    def calc_kpi_scorecard(self) -> pd.DataFrame:
        """KPI 스코어카드 - 벤치마크 대비 현황"""
        results = []
        
        # RC004: YoY Growth
        yoy = self.calc_yoy_growth()
        if len(yoy) > 0:
            latest_yoy = float(yoy.iloc[-1]['yoy_growth_pct'])
            defn = KPI_DEFINITIONS['RC004']
            results.append({
                'code': 'RC004',
                'name_kr': defn.name_kr,
                'category': defn.category.value,
                'actual': latest_yoy,
                'benchmark': defn.benchmark,
                'target': defn.target,
                'unit': defn.unit,
                'status': self._get_status(latest_yoy, defn.benchmark, defn.target, defn.direction)
            })
        
        # OD002: Volatility
        vol = self.calc_volatility()
        defn = KPI_DEFINITIONS['OD002']
        results.append({
            'code': 'OD002',
            'name_kr': defn.name_kr,
            'category': defn.category.value,
            'actual': vol['coefficient_of_variation'],
            'benchmark': defn.benchmark,
            'target': defn.target,
            'unit': defn.unit,
            'status': self._get_status(vol['coefficient_of_variation'], defn.benchmark, defn.target, defn.direction)
        })
        
        # OD001: HHI (Commodity)
        hhi = self.calc_hhi_by_dimension('hs2')
        defn = KPI_DEFINITIONS['OD001']
        results.append({
            'code': 'OD001',
            'name_kr': defn.name_kr + ' (품목)',
            'category': defn.category.value,
            'actual': hhi['hhi'],
            'benchmark': defn.benchmark,
            'target': defn.target,
            'unit': defn.unit,
            'status': self._get_status(hhi['hhi'], defn.benchmark, defn.target, defn.direction)
        })
        
        # RM004: Undervaluation Rate
        underval = self.calc_undervaluation_stats()
        if len(underval) > 0:
            latest_underval = float(underval.iloc[0]['underval_rate'])
            defn = KPI_DEFINITIONS['RM004']
            results.append({
                'code': 'RM004',
                'name_kr': defn.name_kr,
                'category': defn.category.value,
                'actual': latest_underval,
                'benchmark': defn.benchmark,
                'target': defn.target,
                'unit': defn.unit,
                'status': self._get_status(latest_underval, defn.benchmark, defn.target, 'higher')  # 탐지율이므로 higher
            })
        
        # RM005: HS Misclassification
        misclass = self.calc_hs_misclassification_rate()
        if len(misclass) > 0:
            latest_misclass = float(misclass.iloc[0]['misclass_rate'])
            defn = KPI_DEFINITIONS['RM005']
            results.append({
                'code': 'RM005',
                'name_kr': defn.name_kr,
                'category': defn.category.value,
                'actual': latest_misclass,
                'benchmark': defn.benchmark,
                'target': defn.target,
                'unit': defn.unit,
                'status': self._get_status(latest_misclass, defn.benchmark, defn.target, defn.direction)
            })
        
        return pd.DataFrame(results)
    
    def _get_status(self, actual: float, benchmark: float, target: float, direction: str) -> str:
        """KPI 상태 판정"""
        if direction == 'higher':
            if actual >= target:
                return 'Excellent'
            elif actual >= benchmark:
                return 'Good'
            else:
                return 'Needs Improvement'
        else:  # lower is better
            if actual <= target:
                return 'Excellent'
            elif actual <= benchmark:
                return 'Good'
            else:
                return 'Needs Improvement'


# === 유틸리티 함수 ===

def format_currency(value: float, currency: str = 'KRW') -> str:
    """통화 포맷팅"""
    if currency == 'KRW':
        if value >= 1e12:
            return f"₩{value/1e12:.1f}조"
        elif value >= 1e8:
            return f"₩{value/1e8:.0f}억"
        else:
            return f"₩{value:,.0f}"
    elif currency == 'USD':
        if value >= 1e9:
            return f"${value/1e9:.1f}B"
        elif value >= 1e6:
            return f"${value/1e6:.0f}M"
        else:
            return f"${value:,.0f}"
    return f"{value:,.0f}"


def format_percent(value: float, decimals: int = 1) -> str:
    """퍼센트 포맷팅"""
    return f"{value:.{decimals}f}%"


def get_trend_indicator(current: float, previous: float) -> str:
    """추세 인디케이터"""
    if previous == 0:
        return "→"
    change = (current - previous) / previous * 100
    if change > 5:
        return "↑"
    elif change < -5:
        return "↓"
    else:
        return "→"


if __name__ == "__main__":
    # 테스트
    import oracledb
    
    conn = oracledb.connect(
        user="CLRIUSR",
        password="ntancisclri1!",
        dsn="211.239.120.42:3535/NTANCIS"
    )
    
    calc = KPICalculator(conn)
    
    print("=== Executive Summary ===")
    summary = calc.calc_executive_summary()
    for k, v in summary.items():
        print(f"  {k}: {v}")
    
    print("\n=== KPI Scorecard ===")
    scorecard = calc.calc_kpi_scorecard()
    print(scorecard.to_string())
    
    conn.close()
