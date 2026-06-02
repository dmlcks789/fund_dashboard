export const TEMPLATE_FILE_NAME = "성과분석조사_프레임.xlsx";
export const TEMPLATE_URL = `/${encodeURIComponent(TEMPLATE_FILE_NAME)}`;
export const FUND_SHEET_NAME = "펀드 조합현황";
export const INVESTEE_SHEET_NAME = "피투자기업 현황";
export const INVESTEE_DATA_START_ROW = 7; // 0-indexed (8행)
export const INVESTEE_SOURCE_COL_OFFSET = 2; // 소스 col 1-45 → 프레임 col 3-47 고정 오프셋
export const INVESTEE_SECTION_ROW = 5; // 0-indexed = row 6 (섹션 헤더)
export const INVESTEE_SUBHEADER_ROW = 6; // 0-indexed = row 7 (서브 헤더)
export const INVESTEE_FIXED_SRC_COLS = 45; // source 1~45 = 연락처 + DQ코드 (고정 오프셋 구간)
export const TIME_SERIES_SECTION_LABELS = ["매출", "직원수", "기업 인수"] as const;

export const MANAGEMENT_FIELD_LABELS = [
  "DQ01. 운용사명",
  "DQ02. 대표자명",
  "DQ03. 본점 소재지",
  "DQ04. 지점 소재지",
  "DQ05. 운용조합 수(개)",
  "DQ06. 운용조합 규모(AUM)(원)",
  "DQ07. 운용조합 투자잔액(원) *청산 제외",
  "DQ08. 총 임직원 수(명)",
  "DQ09. 심사역 수(명)",
  "DQ10. 부산사무소 상근 임직원 수(해당시)",
  "DQ11. 부산사무소 상근 심사역 수(해당시)"
] as const;

export const ASSOCIATION_FIELD_LABELS = [
  "DQ01. 조합명",
  "DQ02. 조합형태(벤처투자조합, 신기사 등)",
  "DQ03. 펀드결성일",
  "DQ04. 청산일 (해당시)",
  "DQ05. 존속기간",
  "DQ06. 투자기간",
  "DQ07. 약정총액",
  "DQ08. 납입총액",
  "DQ09. 투자자산(원)",
  "DQ10. 미투자자산(원)",
  "DQ11. 배분총액(원)",
  "DQ12. 부산기업 의무투자 금액",
  "DQ13. 부산기업 집행 투자금액"
] as const;

export const INVESTMENT_FIELD_LABELS = [
  "A01. 연도별 투자 기업수",
  "A02. 해당 연도별 투자금액",
  "A03. 연도별 부산기업 투자 기업수(본점만)",
  "A04. 연도별 부산기업 투자금액(본점만)",
  "A05. 투자 후 사업장 유지 부산기업 수",
  "A06. 투자 후 폐업한 부산기업 수",
  "A07. 투자 후 부산으로 이전한 부산 기업 수(본사기준)",
  "A08. 투자 후 본사에서 타 지역으로 이전한 부산기업 수(본사기준)",
  "A09. 해당연도별 회수금액",
  "A10. 해당연도 별 부산기업 회수금액",
  "A11. 향후 회수 예상 금액",
  "A12. 운용펀드 수익률(멀티플) * 영업보고서상 해당연도 조합평가 비율 기재",
  "A13. 청산펀드 수익률(멀티플)",
  "A14. 청산펀드 수익률(IRR)"
] as const;

export const IPO_FIELD_LABELS = [
  "A15. 총 투자기업 중 IPO 기업 수",
  "A16. 총 투자기업 중 부산기업 IPO 기업  수",
  "A17. 본 조합 투자기업 중 타지역 IPO 기업명(서술)"
] as const;

export const INVESTMENT_YEARS = ["2025", "2026", "2027"] as const;
export type InvestmentYear = (typeof INVESTMENT_YEARS)[number];
export type CellPrimitive = string | number | boolean | Date | null;
