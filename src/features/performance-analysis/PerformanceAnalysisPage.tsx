import { ChangeEvent, useEffect, useState } from "react";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";
import "./PerformanceAnalysisPage.css";

const TEMPLATE_FILE_NAME = "성과분석조사_프레임.xlsx";
const TEMPLATE_URL = `/${encodeURIComponent(TEMPLATE_FILE_NAME)}`;
const FUND_SHEET_NAME = "펀드 조합현황";
const INVESTEE_SHEET_NAME = "피투자기업 현황";
const INVESTEE_DATA_START_ROW = 7; // 0-indexed (8행)
const INVESTEE_SOURCE_COL_OFFSET = 2; // 소스 col 1+ → 프레임 col 3+ (운용사명/조합명 2칸 밀림)

const MANAGEMENT_FIELD_LABELS = [
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

const ASSOCIATION_FIELD_LABELS = [
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

const INVESTMENT_FIELD_LABELS = [
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

const IPO_FIELD_LABELS = [
  "A15. 총 투자기업 중 IPO 기업 수",
  "A16. 총 투자기업 중 부산기업 IPO 기업  수",
  "A17. 본 조합 투자기업 중 타지역 IPO 기업명(서술)"
] as const;

const INVESTMENT_YEARS = ["2025", "2026", "2027"] as const;
type InvestmentYear = (typeof INVESTMENT_YEARS)[number];
type CellPrimitive = string | number | boolean | Date | null;

function normalizeLabel(value: string): string {
  return value.toLowerCase().replace(/[^0-9a-z가-힣]/g, "");
}

function getCellText(sheet: XLSX.WorkSheet, row: number, col: number): string {
  const address = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[address];
  if (!cell) {
    return "";
  }
  if (cell.w !== undefined && cell.w !== null) {
    return String(cell.w).trim();
  }
  if (cell.z && cell.v !== undefined && cell.v !== null) {
    try {
      return String(XLSX.SSF.format(cell.z, cell.v)).trim();
    } catch {
      // fall through to raw value text
    }
  }
  return String(cell.v ?? "").trim();
}

function getCellRawValue(sheet: XLSX.WorkSheet, row: number, col: number): CellPrimitive {
  const address = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[address];
  if (!cell || cell.v === undefined || cell.v === null) {
    return null;
  }
  if (typeof cell.v === "number" || typeof cell.v === "boolean" || cell.v instanceof Date) {
    return cell.v;
  }
  const text = String(cell.v).trim();
  return text === "" ? null : text;
}

function formatNumberForDisplay(value: number): string {
  return new Intl.NumberFormat("en-US", {
    useGrouping: false,
    maximumFractionDigits: 20
  }).format(value);
}

function isBlankValue(value: CellPrimitive): boolean {
  return value === null || (typeof value === "string" && value.trim() === "");
}

function labelsMatch(a: string, b: string): boolean {
  const na = normalizeLabel(a);
  const nb = normalizeLabel(b);
  return na === nb || na.includes(nb) || nb.includes(na);
}

function findCellByLabel(
  sheet: XLSX.WorkSheet,
  targetLabel: string,
  rowRange?: { startRow: number; endRow: number }
): { row: number; col: number } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const startRow = rowRange?.startRow ?? range.s.r;
  const endRow = rowRange?.endRow ?? range.e.r;

  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (text && labelsMatch(text, targetLabel)) {
        return { row, col };
      }
    }
  }
  return null;
}

function findSheetNameContainingLabel(workbook: XLSX.WorkBook, targetLabel: string): string | null {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      continue;
    }
    if (findCellByLabel(sheet, targetLabel)) {
      return sheetName;
    }
  }
  return null;
}

function findSheetNameByPreferredOrder(workbook: XLSX.WorkBook): string | null {
  return workbook.SheetNames.find((name) => labelsMatch(name, FUND_SHEET_NAME)) ?? workbook.SheetNames[0] ?? null;
}

function extractCaseIndexFromText(text: string): number | null {
  const compact = text.replace(/\s/g, "");
  const match = compact.match(/수범기업[#:\-]?(1|2|3)/);
  if (!match) {
    return null;
  }
  const index = Number(match[1]);
  return index >= 1 && index <= 3 ? index : null;
}

function findCaseMarkersByPattern(sheet: XLSX.WorkSheet): Array<{ row: number; col: number } | null> {
  const markers: Array<{ row: number; col: number } | null> = [null, null, null];
  const ref = sheet["!ref"];
  if (!ref) {
    return markers;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text) {
        continue;
      }
      const index = extractCaseIndexFromText(text);
      if (!index) {
        continue;
      }
      if (!markers[index - 1]) {
        markers[index - 1] = { row, col };
      }
    }
  }
  return markers;
}

function findColumnByHeaderLabel(sheet: XLSX.WorkSheet, headerLabel: string): number | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const headerRow = 2;
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const text = getCellText(sheet, headerRow, col);
    if (text && labelsMatch(text, headerLabel)) {
      return col;
    }
  }
  return null;
}

function findColumnByHeaderCode(sheet: XLSX.WorkSheet, fieldCode: string): number | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const headerRow = 2;
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const text = getCellText(sheet, headerRow, col);
    if (extractFieldCode(text) === fieldCode) {
      return col;
    }
  }
  return null;
}

function findColumnByHeaderCodeOccurrence(
  sheet: XLSX.WorkSheet,
  fieldCode: string,
  occurrenceIndex: number
): number | null {
  const ref = sheet["!ref"];
  if (!ref || occurrenceIndex < 0) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const headerRow = 2;
  let matchedIndex = 0;
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const text = getCellText(sheet, headerRow, col);
    if (extractFieldCode(text) !== fieldCode) {
      continue;
    }
    if (matchedIndex === occurrenceIndex) {
      return col;
    }
    matchedIndex += 1;
  }
  return null;
}

function findInvestmentColumnByYearAndCode(sheet: XLSX.WorkSheet, year: InvestmentYear, fieldCode: string): number | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const yearRow = 1;
  const headerRow = 2;
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const yearText = getCellText(sheet, yearRow, col);
    const headerText = getCellText(sheet, headerRow, col);
    if (isStrictYearToken(yearText, year) && extractFieldCode(headerText) === fieldCode) {
      return col;
    }
  }
  return null;
}

function findRowToWrite(sheet: XLSX.WorkSheet): number {
  const ref = sheet["!ref"];
  if (!ref) {
    return 3;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let row = 3; row <= range.e.r; row += 1) {
    let hasData = false;
    for (let col = 1; col <= range.e.c; col += 1) {
      if (getCellText(sheet, row, col)) {
        hasData = true;
        break;
      }
    }
    if (!hasData) {
      return row;
    }
  }
  return range.e.r + 1;
}

type ParsedCase = {
  company: string;
  content: string;
};

function parseBestPracticeCases(sourceSheet: XLSX.WorkSheet): ParsedCase[] {
  const markers = findCaseMarkersByPattern(sourceSheet);
  const sourceRef = sourceSheet["!ref"] ? XLSX.utils.decode_range(sourceSheet["!ref"]) : null;
  if (!sourceRef) {
    return [];
  }

  const results: ParsedCase[] = [];
  for (let index = 0; index < 3; index += 1) {
    const marker = markers[index];
    if (!marker) {
      results.push({ company: "", content: "" });
      continue;
    }

    const nextMarker = markers[index + 1];
    const sectionEndRow = nextMarker && nextMarker.row > marker.row ? nextMarker.row : sourceRef.e.r + 1;

    const companyLabel =
      findLabelInRowWindow(sourceSheet, marker.row + 1, marker.col, marker.col + 24, "기업명") ??
      findCellByLabel(sourceSheet, "기업명", {
        startRow: marker.row,
        endRow: Math.min(marker.row + 4, sectionEndRow - 1)
      });
    const companyCell = companyLabel
      ? findFirstMeaningfulValueRightOfCell(sourceSheet, companyLabel.row, companyLabel.col, ["기업명", "내용"])
      : null;

    let content = "";
    for (let row = marker.row + 2; row < sectionEndRow; row += 1) {
      const cell = findFirstNonEmptyCellInRow(sourceSheet, row, marker.col);
      if (!cell) {
        continue;
      }
      if (labelsMatch(cell.value, "기업명") || labelsMatch(cell.value, "내용")) {
        continue;
      }
      content = cell.value;
      break;
    }

    results.push({ company: companyCell?.value ?? "", content });
  }

  return results;
}

function extractFieldCode(label: string): string {
  const normalized = label.trim().toUpperCase();
  const matched = normalized.match(/^([A-Z]{1,2}\d{2})\s*\./);
  return matched?.[1] ?? "";
}

function findFirstValueRightOfCell(sheet: XLSX.WorkSheet, row: number, col: number): { col: number; value: string } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let currentCol = col + 1; currentCol <= range.e.c; currentCol += 1) {
    const value = getCellText(sheet, row, currentCol);
    if (value) {
      return { col: currentCol, value };
    }
  }
  return null;
}

function findFirstMeaningfulValueRightOfCell(
  sheet: XLSX.WorkSheet,
  row: number,
  col: number,
  ignoredLabels: string[]
): { col: number; value: string } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let currentCol = col + 1; currentCol <= range.e.c; currentCol += 1) {
    const value = getCellText(sheet, row, currentCol);
    if (!value) {
      continue;
    }
    if (ignoredLabels.some((label) => labelsMatch(value, label))) {
      continue;
    }
    return { col: currentCol, value };
  }
  return null;
}

function findFirstNonEmptyCellInRow(
  sheet: XLSX.WorkSheet,
  row: number,
  startCol: number
): { col: number; value: string } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let col = Math.max(startCol, range.s.c); col <= range.e.c; col += 1) {
    const value = getCellText(sheet, row, col);
    if (value) {
      return { col, value };
    }
  }
  return null;
}

function findLabelInRowWindow(
  sheet: XLSX.WorkSheet,
  row: number,
  startCol: number,
  endCol: number,
  targetLabel: string
): { row: number; col: number } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const r = Math.max(range.s.r, Math.min(range.e.r, row));
  const fromCol = Math.max(range.s.c, startCol);
  const toCol = Math.min(range.e.c, endCol);
  for (let col = fromCol; col <= toCol; col += 1) {
    const text = getCellText(sheet, r, col);
    if (text && labelsMatch(text, targetLabel)) {
      return { row: r, col };
    }
  }
  return null;
}

function findAllValuesRightOfColumnALabel(sheet: XLSX.WorkSheet, targetLabel: string): string {
  const ref = sheet["!ref"];
  if (!ref) return "";
  const range = XLSX.utils.decode_range(ref);
  const values: string[] = [];
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const label = getCellText(sheet, row, 0);
    if (!label || !labelsMatch(label, targetLabel)) continue;
    for (let col = 1; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (text) {
        values.push(text);
        break;
      }
    }
  }
  return values.join(", ");
}

function findValueRightOfColumnALabel(sheet: XLSX.WorkSheet, targetLabel: string): CellPrimitive {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const label = getCellText(sheet, row, 0);
    if (!label || !labelsMatch(label, targetLabel)) {
      continue;
    }
    for (let col = 1; col <= range.e.c; col += 1) {
      const rawValue = getCellRawValue(sheet, row, col);
      const textValue = getCellText(sheet, row, col);
      if (!isBlankValue(rawValue) || textValue !== "") {
        if (textValue !== "") {
          return textValue;
        }
        return rawValue;
      }
    }
    return null;
  }
  return null;
}

function findValueRightOfColumnAFieldCodeOccurrence(
  sheet: XLSX.WorkSheet,
  fieldCode: string,
  occurrenceIndex: number
): CellPrimitive {
  if (!fieldCode || occurrenceIndex < 0) {
    return null;
  }
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  let matchedIndex = 0;

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const label = getCellText(sheet, row, 0);
    if (!label || extractFieldCode(label) !== fieldCode) {
      continue;
    }
    if (matchedIndex !== occurrenceIndex) {
      matchedIndex += 1;
      continue;
    }
    for (let col = 1; col <= range.e.c; col += 1) {
      const rawValue = getCellRawValue(sheet, row, col);
      const textValue = getCellText(sheet, row, col);
      if (!isBlankValue(rawValue) || textValue !== "") {
        return textValue !== "" ? textValue : rawValue;
      }
    }
    return null;
  }
  return null;
}

function getFieldCodeNumber(fieldCode: string): number | null {
  const matched = fieldCode.match(/^[A-Z]{1,2}(\d{2})$/);
  if (!matched) {
    return null;
  }
  return Number(matched[1]);
}

function getAssociationOccurrenceIndex(fieldCode: string): number {
  const fieldNumber = getFieldCodeNumber(fieldCode);
  if (fieldCode.startsWith("DQ") && fieldNumber !== null && fieldNumber <= 11) {
    return 1;
  }
  return 0;
}

function applyLabelMappedSection(
  sourceSheet: XLSX.WorkSheet,
  targetSheet: XLSX.WorkSheet,
  writeRow: number,
  labels: readonly string[]
) {
  for (const fieldLabel of labels) {
    const value = findValueRightOfColumnALabel(sourceSheet, fieldLabel);
    if (isBlankValue(value)) {
      continue;
    }
    const targetCol = findColumnByHeaderLabel(targetSheet, fieldLabel);
    if (targetCol === null) {
      continue;
    }
    setSheetCellText(targetSheet, writeRow, targetCol, value);
  }
}

function findRowByColumnALabel(sheet: XLSX.WorkSheet, targetLabel: string): number | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const label = getCellText(sheet, row, 0);
    if (label && labelsMatch(label, targetLabel)) {
      return row;
    }
  }
  return null;
}

function findRowByColumnAFieldCode(sheet: XLSX.WorkSheet, fieldCode: string): number | null {
  if (!fieldCode) {
    return null;
  }
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const label = getCellText(sheet, row, 0);
    if (!label) {
      continue;
    }
    if (extractFieldCode(label) === fieldCode) {
      return row;
    }
  }
  return null;
}

function isStrictYearToken(text: string, year: InvestmentYear): boolean {
  const compact = text.replace(/\s/g, "");
  const matchedYear = compact.match(/(2025|2026|2027)/)?.[1];
  return matchedYear === year;
}

function findYearColumnsNearAnchorRow(
  sheet: XLSX.WorkSheet,
  anchorRow: number,
  years: readonly InvestmentYear[]
): Partial<Record<InvestmentYear, number>> {
  const ref = sheet["!ref"];
  if (!ref) {
    return {};
  }

  const range = XLSX.utils.decode_range(ref);
  const startRow = Math.max(range.s.r, anchorRow - 30);
  const endRow = Math.min(range.e.r, anchorRow + 30);
  let bestMatch: Partial<Record<InvestmentYear, number>> = {};
  let bestDistance = Number.POSITIVE_INFINITY;

  for (let row = startRow; row <= endRow; row += 1) {
    const current: Partial<Record<InvestmentYear, number>> = {};
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const text = getCellText(sheet, row, col);
        if (!text) {
          continue;
        }
        for (const year of years) {
          if (isStrictYearToken(text, year)) {
            current[year] = col;
          }
        }
      }

    const currentCount = Object.keys(current).length;
    const bestCount = Object.keys(bestMatch).length;
    const distance = Math.abs(row - anchorRow);
    if (
      currentCount > bestCount ||
      (currentCount === bestCount && currentCount > 0 && distance < bestDistance)
    ) {
      bestMatch = current;
      bestDistance = distance;
    }
  }

  return bestMatch;
}

function findYearColumnsInSingleRow(
  sheet: XLSX.WorkSheet,
  years: readonly InvestmentYear[]
): { row: number; cols: Record<InvestmentYear, number> } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const cols: Partial<Record<InvestmentYear, number>> = {};
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text) {
        continue;
      }
      for (const year of years) {
        if (cols[year] === undefined && isStrictYearToken(text, year)) {
          cols[year] = col;
        }
      }
    }
    if (years.every((year) => cols[year] !== undefined)) {
      return { row, cols: cols as Record<InvestmentYear, number> };
    }
  }
  return null;
}

function findYearColumnsAroundRow(
  sheet: XLSX.WorkSheet,
  centerRow: number,
  years: readonly InvestmentYear[],
  windowSize = 3
): { row: number; cols: Record<InvestmentYear, number> } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const startRow = Math.max(range.s.r, centerRow - windowSize);
  const endRow = Math.min(range.e.r, centerRow + windowSize);

  let best: { row: number; cols: Partial<Record<InvestmentYear, number>>; count: number; distance: number } | null = null;

  for (let row = startRow; row <= endRow; row += 1) {
    const cols: Partial<Record<InvestmentYear, number>> = {};
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text) {
        continue;
      }
      for (const year of years) {
        if (cols[year] === undefined && isStrictYearToken(text, year)) {
          cols[year] = col;
        }
      }
    }

    const count = years.filter((year) => cols[year] !== undefined).length;
    if (count === 0) {
      continue;
    }
    const distance = Math.abs(row - centerRow);
    if (!best || count > best.count || (count === best.count && distance < best.distance)) {
      best = { row, cols, count, distance };
    }
  }

  if (!best || !years.every((year) => best.cols[year] !== undefined)) {
    return null;
  }

  return {
    row: best.row,
    cols: best.cols as Record<InvestmentYear, number>
  };
}

function resolveInvestmentTargetColumn(
  sheet: XLSX.WorkSheet,
  headerRow: number,
  yearStartCol: number,
  fieldCode: string,
  codeIndex: number
): number | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const candidate = yearStartCol + codeIndex;
  if (candidate <= range.e.c && extractFieldCode(getCellText(sheet, headerRow, candidate)) === fieldCode) {
    return candidate;
  }
  const searchEndCol = Math.min(range.e.c, yearStartCol + 40);
  for (let col = yearStartCol; col <= searchEndCol; col += 1) {
    if (extractFieldCode(getCellText(sheet, headerRow, col)) === fieldCode) {
      return col;
    }
  }
  return null;
}

function findYearCellNearRow(
  sheet: XLSX.WorkSheet,
  year: InvestmentYear,
  centerRow: number,
  windowSize = 6
): { row: number; col: number } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const startRow = Math.max(range.s.r, centerRow - windowSize);
  const endRow = Math.min(range.e.r, centerRow + windowSize);

  let best: { row: number; col: number; distance: number } | null = null;
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text || !isStrictYearToken(text, year)) {
        continue;
      }
      const distance = Math.abs(row - centerRow);
      if (!best || distance < best.distance) {
        best = { row, col, distance };
      }
    }
  }

  return best ? { row: best.row, col: best.col } : null;
}

function findInvestmentTargetLayout(
  sheet: XLSX.WorkSheet,
  years: readonly InvestmentYear[]
): { headerRow: number; yearStartCols: Record<InvestmentYear, number> } | null {
  const ref = sheet["!ref"];
  if (!ref) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);

  const repeatedA01Cols: number[] = [];
  const repeatedHeaderRow = 2;
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    if (extractFieldCode(getCellText(sheet, repeatedHeaderRow, col)) === "A01") {
      repeatedA01Cols.push(col);
    }
  }
  if (repeatedA01Cols.length >= years.length) {
    const yearStartCols = {} as Record<InvestmentYear, number>;
    years.forEach((year, index) => {
      yearStartCols[year] = repeatedA01Cols[index];
    });
    return { headerRow: repeatedHeaderRow, yearStartCols };
  }

  const singleYearRow = findYearColumnsInSingleRow(sheet, years);
  if (singleYearRow) {
    return { headerRow: singleYearRow.row + 1, yearStartCols: singleYearRow.cols };
  }

  return null;
}

function findCellByFieldCodeInColumnRange(
  sheet: XLSX.WorkSheet,
  fieldCode: string,
  startCol: number,
  endCol: number
): { row: number; col: number } | null {
  const ref = sheet["!ref"];
  if (!ref || !fieldCode) {
    return null;
  }
  const range = XLSX.utils.decode_range(ref);
  const fromCol = Math.max(range.s.c, startCol);
  const toCol = Math.min(range.e.c, endCol);
  if (fromCol > toCol) {
    return null;
  }

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = fromCol; col <= toCol; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text) {
        continue;
      }
      if (extractFieldCode(text) === fieldCode) {
        return { row, col };
      }
    }
  }
  return null;
}

function buildYearColumnRanges(
  yearColumns: Partial<Record<InvestmentYear, number>>,
  fallbackEndCol: number
): Partial<Record<InvestmentYear, { startCol: number; endCol: number }>> {
  const ordered = INVESTMENT_YEARS
    .map((year) => ({ year, col: yearColumns[year] }))
    .filter((entry): entry is { year: InvestmentYear; col: number } => entry.col !== undefined)
    .sort((a, b) => a.col - b.col);

  const result: Partial<Record<InvestmentYear, { startCol: number; endCol: number }>> = {};
  for (let index = 0; index < ordered.length; index += 1) {
    const current = ordered[index];
    const next = ordered[index + 1];
    result[current.year] = {
      startCol: current.col,
      endCol: next ? next.col - 1 : fallbackEndCol
    };
  }
  return result;
}

function isLikelyHorizontalHeaderRow(sheet: XLSX.WorkSheet, row: number, col: number): boolean {
  let codeNeighborCount = 0;
  for (let offset = -2; offset <= 2; offset += 1) {
    if (offset === 0) {
      continue;
    }
    const text = getCellText(sheet, row, Math.max(0, col + offset));
    if (/^(dq|a)\s*0*\d+/i.test(text.trim())) {
      codeNeighborCount += 1;
    }
  }
  return codeNeighborCount >= 1;
}

function findWriteCellRightOfLabel(sheet: XLSX.WorkSheet, labelRow: number, labelCol: number): { row: number; col: number } {
  const ref = sheet["!ref"];
  const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { r: labelRow, c: labelCol + 1 } };
  const merges = sheet["!merges"] ?? [];

  for (let col = labelCol + 1; col <= range.e.c; col += 1) {
    const merge = merges.find(
      (m) => m.s.r <= labelRow && m.e.r >= labelRow && m.s.c <= col && m.e.c >= col
    );
    if (merge) {
      if (merge.s.r !== labelRow || merge.s.c <= labelCol) {
        continue;
      }
      return { row: merge.s.r, col: merge.s.c };
    }
    return { row: labelRow, col };
  }

  return { row: labelRow, col: labelCol + 1 };
}

function setSheetCellText(sheet: XLSX.WorkSheet, row: number, col: number, value: CellPrimitive) {
  const address = XLSX.utils.encode_cell({ r: row, c: col });
  const prev = sheet[address];
  if (isBlankValue(value)) {
    delete sheet[address];
  } else if (typeof value === "number") {
    sheet[address] = {
      ...(prev ?? {}),
      t: "n",
      v: value,
      w: formatNumberForDisplay(value)
    };
  } else if (typeof value === "boolean") {
    sheet[address] = {
      ...(prev ?? {}),
      t: "b",
      v: value
    };
  } else if (value instanceof Date) {
    sheet[address] = {
      ...(prev ?? {}),
      t: "d",
      v: value
    };
  } else {
    sheet[address] = {
      ...(prev ?? {}),
      t: "s",
      v: value,
      w: value
    };
  }

  const ref = sheet["!ref"];
  if (!ref) {
    sheet["!ref"] = XLSX.utils.encode_range({ s: { r: row, c: col }, e: { r: row, c: col } });
    return;
  }
  const range = XLSX.utils.decode_range(ref);
  if (row < range.s.r) {
    range.s.r = row;
  }
  if (col < range.s.c) {
    range.s.c = col;
  }
  if (row > range.e.r) {
    range.e.r = row;
  }
  if (col > range.e.c) {
    range.e.c = col;
  }
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function writeValueAtLabelTarget(sheet: XLSX.WorkSheet, labelCell: { row: number; col: number }, value: CellPrimitive) {
  if (isLikelyHorizontalHeaderRow(sheet, labelCell.row, labelCell.col)) {
    setSheetCellText(sheet, labelCell.row + 1, labelCell.col, value);
    return;
  }
  const writeCell = findWriteCellRightOfLabel(sheet, labelCell.row, labelCell.col);
  setSheetCellText(sheet, writeCell.row, writeCell.col, value);
}

function buildSheetHtml(sheetName: string, sheet: XLSX.WorkSheet): string {
  const tableHtml = XLSX.utils.sheet_to_html(sheet, { id: `sheet-${sheetName}` });
  return `<!doctype html>
<html lang="ko">
  <head>
    <meta charset="utf-8" />
    <style>
      body {
        margin: 0;
        padding: 10px;
        background: #fff;
        color: #0f172a;
        font-family: "Malgun Gothic", "Noto Sans KR", sans-serif;
      }
      table { border-collapse: collapse; }
      td, th {
        border: 1px solid #cbd5e1;
        padding: 6px 8px;
        font-size: 13px;
        white-space: nowrap;
        vertical-align: middle;
      }
    </style>
  </head>
  <body>${tableHtml}</body>
</html>`;
}

function cloneWorkbook(workbook: XLSX.WorkBook): XLSX.WorkBook {
  const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx", cellStyles: true });
  return XLSX.read(bytes, { type: "array", cellStyles: true });
}

function getCellValueForCompare(sheet: XLSX.WorkSheet | undefined, row: number, col: number): CellPrimitive {
  if (!sheet) {
    return null;
  }
  const raw = getCellRawValue(sheet, row, col);
  if (!isBlankValue(raw)) {
    return raw;
  }
  const text = getCellText(sheet, row, col);
  return text === "" ? null : text;
}

function normalizeComparableValue(value: CellPrimitive): string {
  if (value === null) {
    return "null";
  }
  if (value instanceof Date) {
    return `date:${value.toISOString()}`;
  }
  return `${typeof value}:${String(value)}`;
}

function getWorkbookValueChanges(current: XLSX.WorkBook, base: XLSX.WorkBook) {
  const changes: Array<{ sheetName: string; row: number; col: number; value: CellPrimitive }> = [];
  for (const sheetName of current.SheetNames) {
    const currentSheet = current.Sheets[sheetName];
    if (!currentSheet) {
      continue;
    }
    const baseSheet = base.Sheets[sheetName];
    const currentRef = currentSheet["!ref"] ? XLSX.utils.decode_range(currentSheet["!ref"]) : null;
    const baseRef = baseSheet?.["!ref"] ? XLSX.utils.decode_range(baseSheet["!ref"]) : null;
    if (!currentRef && !baseRef) {
      continue;
    }
    const maxRow = Math.max(currentRef?.e.r ?? 0, baseRef?.e.r ?? 0);
    const maxCol = Math.max(currentRef?.e.c ?? 0, baseRef?.e.c ?? 0);
    for (let row = 0; row <= maxRow; row += 1) {
      for (let col = 0; col <= maxCol; col += 1) {
        const currentValue = getCellValueForCompare(currentSheet, row, col);
        const baseValue = getCellValueForCompare(baseSheet, row, col);
        if (normalizeComparableValue(currentValue) === normalizeComparableValue(baseValue)) {
          continue;
        }
        changes.push({ sheetName, row, col, value: currentValue });
      }
    }
  }
  return changes;
}

function coercePrimitiveForExcel(value: CellPrimitive): string | number | boolean | Date | null {
  if (value === null || typeof value === "number" || typeof value === "boolean" || value instanceof Date) {
    return value;
  }
  const trimmed = value.trim();
  if (trimmed === "") {
    return null;
  }
  return trimmed;
}

function extendSheetStyles(worksheet: ExcelJS.Worksheet, firstDataRow: number): void {
  const lastRowNum = worksheet.lastRow?.number;
  if (!lastRowNum || lastRowNum <= firstDataRow) return;

  const srcRow = worksheet.getRow(firstDataRow);

  for (let r = firstDataRow + 1; r <= lastRowNum; r++) {
    const row = worksheet.getRow(r);

    // 이미 템플릿 스타일(테두리)이 있으면 스킵
    let hasStyle = false;
    for (let c = 1; c <= 8; c++) {
      if (row.getCell(c).border) {
        hasStyle = true;
        break;
      }
    }
    if (hasStyle) continue;

    // 완전히 빈 행은 스킵
    let hasValue = false;
    row.eachCell((cell) => {
      if (cell.value != null && cell.value !== "") hasValue = true;
    });
    if (!hasValue) continue;

    // 첫 데이터 행의 스타일 복사
    const colCount = worksheet.columnCount || 72;
    for (let c = 1; c <= colCount; c++) {
      const srcCell = srcRow.getCell(c);
      const destCell = row.getCell(c);
      if (srcCell.border) destCell.border = JSON.parse(JSON.stringify(srcCell.border)) as ExcelJS.Borders;
      if (srcCell.numFmt) destCell.numFmt = srcCell.numFmt;
      if (srcCell.alignment) destCell.alignment = JSON.parse(JSON.stringify(srcCell.alignment)) as ExcelJS.Alignment;
    }
  }

  // 데이터 유효성 확장: ExcelJS는 셀별(T8, T9...) 저장 방식 → 첫 행 정의를 신규 행에 복사
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = worksheet as any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const model = ws.dataValidations?.model as Record<string, any> | undefined;
  if (model) {
    // 첫 데이터 행의 유효성 정의 수집 (예: "T8", "O8" 등)
    const firstRowDvs: Array<{ col: string; dv: unknown }> = [];
    for (const [ref, dv] of Object.entries(model)) {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if (m && parseInt(m[2]) === firstDataRow) {
        firstRowDvs.push({ col: m[1], dv });
      }
    }

    // 이미 유효성이 있는 행 목록
    const existingRows = new Set<number>();
    for (const ref of Object.keys(model)) {
      const m = ref.match(/^[A-Z]+(\d+)$/);
      if (m) existingRows.add(parseInt(m[1]));
    }

    // 신규 행에 유효성 복사
    for (let r = firstDataRow + 1; r <= lastRowNum; r++) {
      if (existingRows.has(r)) continue;
      const row = worksheet.getRow(r);
      let hasValue = false;
      row.eachCell((cell) => {
        if (cell.value != null && cell.value !== "") hasValue = true;
      });
      if (!hasValue) continue;
      for (const { col, dv } of firstRowDvs) {
        model[`${col}${r}`] = { ...(dv as object) };
      }
    }
  }
}

function findInvesteeWriteRow(sheet: XLSX.WorkSheet): number {
  const ref = sheet["!ref"];
  if (!ref) return INVESTEE_DATA_START_ROW;
  const range = XLSX.utils.decode_range(ref);
  for (let row = INVESTEE_DATA_START_ROW; row <= range.e.r; row += 1) {
    let hasData = false;
    for (let col = 1; col <= range.e.c; col += 1) {
      if (getCellText(sheet, row, col)) {
        hasData = true;
        break;
      }
    }
    if (!hasData) return row;
  }
  return range.e.r + 1;
}

function applyInvesteeFirmsToTemplate(sourceWorkbook: XLSX.WorkBook, targetSheet: XLSX.WorkSheet) {
  const sourceSheet = sourceWorkbook.Sheets[INVESTEE_SHEET_NAME];
  if (!sourceSheet) return;

  const fundSheetName =
    FUND_SHEET_NAME in sourceWorkbook.Sheets
      ? FUND_SHEET_NAME
      : findSheetNameContainingLabel(sourceWorkbook, "DQ01. 운용사명");
  const fundSheet = fundSheetName ? sourceWorkbook.Sheets[fundSheetName] : null;

  const 운용사명Raw = fundSheet ? findAllValuesRightOfColumnALabel(fundSheet, "DQ01. 운용사명") : "";
  const 운용사명: CellPrimitive = 운용사명Raw || null;
  const 조합명 = fundSheet ? findValueRightOfColumnALabel(fundSheet, "DQ01. 조합명") : null;

  const ref = sourceSheet["!ref"];
  if (!ref) return;
  const range = XLSX.utils.decode_range(ref);

  for (let srcRow = INVESTEE_DATA_START_ROW; srcRow <= range.e.r; srcRow += 1) {
    let hasContent = false;
    for (let col = 1; col <= range.e.c; col += 1) {
      if (getCellText(sourceSheet, srcRow, col)) {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) continue;

    const writeRow = findInvesteeWriteRow(targetSheet);
    setSheetCellText(targetSheet, writeRow, 0, writeRow - INVESTEE_DATA_START_ROW + 1);
    if (!isBlankValue(운용사명)) setSheetCellText(targetSheet, writeRow, 1, 운용사명);
    if (!isBlankValue(조합명)) setSheetCellText(targetSheet, writeRow, 2, 조합명);

    for (let srcCol = 1; srcCol <= range.e.c; srcCol += 1) {
      const rawValue = getCellRawValue(sourceSheet, srcRow, srcCol);
      const textValue = getCellText(sourceSheet, srcRow, srcCol);
      if (isBlankValue(rawValue) && textValue === "") continue;
      const value = textValue !== "" ? textValue : rawValue;
      setSheetCellText(targetSheet, writeRow, srcCol + INVESTEE_SOURCE_COL_OFFSET, value);
    }
  }
}

export default function PerformanceAnalysisPage() {
  const [isLoading, setIsLoading] = useState(true);
  const [errorMessage, setErrorMessage] = useState("");
  const [templateWorkbook, setTemplateWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheetName, setSelectedSheetName] = useState("");
  const [sheetHtml, setSheetHtml] = useState("");
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [isApplying, setIsApplying] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [unmatchedFiles, setUnmatchedFiles] = useState<string[]>([]);
  const [showUnmatched, setShowUnmatched] = useState(false);

  const downloadWorkbook = async () => {
    if (!workbook || !templateWorkbook) {
      return;
    }
    setIsDownloading(true);
    setErrorMessage("");
    try {
      const response = await fetch(TEMPLATE_URL, { cache: "no-store" });
      if (!response.ok) {
        setErrorMessage(`파일을 찾을 수 없습니다: ${TEMPLATE_FILE_NAME}`);
        return;
      }
      const templateBuffer = await response.arrayBuffer();
      const excelWorkbook = new ExcelJS.Workbook();
      await excelWorkbook.xlsx.load(templateBuffer);

      const changes = getWorkbookValueChanges(workbook, templateWorkbook);
      for (const change of changes) {
        const worksheet = excelWorkbook.getWorksheet(change.sheetName);
        if (!worksheet) {
          continue;
        }
        const cell = worksheet.getCell(change.row + 1, change.col + 1);
        cell.value = coercePrimitiveForExcel(change.value);
      }

      // 신규 행에 테두리/콤보박스 스타일 확장
      const investeeWS = excelWorkbook.getWorksheet(INVESTEE_SHEET_NAME);
      if (investeeWS) extendSheetStyles(investeeWS, INVESTEE_DATA_START_ROW + 1);
      const fundWS = excelWorkbook.getWorksheet(FUND_SHEET_NAME);
      if (fundWS) extendSheetStyles(fundWS, 4); // 펀드 조합현황 첫 데이터 행 = 4 (1-indexed)

      const outputBuffer = await excelWorkbook.xlsx.writeBuffer();
      const blob = new Blob([outputBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
      const now = new Date();
      const pad = (n: number) => String(n).padStart(2, "0");
      const timestamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
      const filename = `성과분석조사_프레임_취합_${timestamp}.xlsx`;

      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(link.href);
    } catch (error) {
      console.error(error);
      setErrorMessage("다운로드 파일 생성 중 오류가 발생했습니다.");
    } finally {
      setIsDownloading(false);
    }
  };

  const selectSheet = (nextWorkbook: XLSX.WorkBook, sheetName: string) => {
    const sheet = nextWorkbook.Sheets[sheetName];
    if (!sheet) {
      setSheetHtml("");
      return;
    }
    setSelectedSheetName(sheetName);
    setSheetHtml(buildSheetHtml(sheetName, sheet));
  };

  const loadTemplateWorkbook = async () => {
    setIsLoading(true);
    setErrorMessage("");
    try {
      const response = await fetch(TEMPLATE_URL, { cache: "no-store" });
      if (!response.ok) {
        setErrorMessage(`파일을 찾을 수 없습니다: ${TEMPLATE_FILE_NAME}`);
        setTemplateWorkbook(null);
        setWorkbook(null);
        setSheetNames([]);
        setSelectedSheetName("");
        setSheetHtml("");
        return;
      }

      const buffer = await response.arrayBuffer();
      const loadedTemplate = XLSX.read(buffer, { type: "array", cellStyles: true });
      const initialWorkbook = cloneWorkbook(loadedTemplate);
      const names = initialWorkbook.SheetNames;

      setTemplateWorkbook(loadedTemplate);
      setWorkbook(initialWorkbook);
      setSheetNames(names);

      if (names.length > 0) {
        const defaultSheet = names.find((name) => labelsMatch(name, FUND_SHEET_NAME)) ?? names[0];
        selectSheet(initialWorkbook, defaultSheet);
      } else {
        setErrorMessage("시트를 찾지 못했습니다.");
        setSheetHtml("");
      }
    } catch (error) {
      console.error(error);
      setErrorMessage("엑셀 파일을 불러오는 중 오류가 발생했습니다.");
      setTemplateWorkbook(null);
      setWorkbook(null);
      setSheetNames([]);
      setSelectedSheetName("");
      setSheetHtml("");
    } finally {
      setIsLoading(false);
    }
  };

  const onUploadFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setUploadedFiles(Array.from(event.target.files ?? []));
  };

  const applyManagerInfoToTemplate = async () => {
    if (uploadedFiles.length === 0 || !templateWorkbook) {
      return;
    }

    setIsApplying(true);
    setErrorMessage("");
    setUnmatchedFiles([]);
    setShowUnmatched(false);
    try {
      const workingWorkbook = cloneWorkbook(templateWorkbook);
      const targetSheet = workingWorkbook.Sheets[FUND_SHEET_NAME];
      if (!targetSheet) {
        setErrorMessage("템플릿에서 펀드 조합현황 시트를 찾지 못했습니다.");
        return;
      }
      const targetInvesteeSheet = workingWorkbook.Sheets[INVESTEE_SHEET_NAME];

      const unmatched: string[] = [];

      for (const sourceFile of uploadedFiles) {
        const sourceBuffer = await sourceFile.arrayBuffer();
        const sourceWorkbook = XLSX.read(sourceBuffer, {
          type: "array",
          cellNF: true,
          cellStyles: true,
          cellDates: true,
          dateNF: "yyyy-mm-dd"
        });

        const hasFundSheet = FUND_SHEET_NAME in sourceWorkbook.Sheets;
        const hasInvesteeSheet = INVESTEE_SHEET_NAME in sourceWorkbook.Sheets;

        if (!hasFundSheet || !hasInvesteeSheet) {
          unmatched.push(sourceFile.name);
          continue;
        }

        if (hasFundSheet) {
          const sourceSheet = sourceWorkbook.Sheets[FUND_SHEET_NAME]!;

          const writeRow = findRowToWrite(targetSheet);
          setSheetCellText(targetSheet, writeRow, 0, writeRow - 2);

          applyLabelMappedSection(sourceSheet, targetSheet, writeRow, MANAGEMENT_FIELD_LABELS);
          applyLabelMappedSection(sourceSheet, targetSheet, writeRow, ASSOCIATION_FIELD_LABELS);
          applyLabelMappedSection(sourceSheet, targetSheet, writeRow, IPO_FIELD_LABELS);

          const targetInvestmentLayout = findInvestmentTargetLayout(targetSheet, INVESTMENT_YEARS);
          const sourceA01Row = findRowByColumnAFieldCode(sourceSheet, "A01");
          if (sourceA01Row !== null && targetInvestmentLayout) {
            const targetHeaderRow = targetInvestmentLayout.headerRow;
            for (const year of INVESTMENT_YEARS) {
              const sourceYearCell = findYearCellNearRow(sourceSheet, year, sourceA01Row - 1, 6);
              if (!sourceYearCell) {
                continue;
              }
              const yearStartCol = targetInvestmentLayout.yearStartCols[year];
              for (let codeIndex = 0; codeIndex < INVESTMENT_FIELD_LABELS.length; codeIndex += 1) {
                const fieldCode = extractFieldCode(INVESTMENT_FIELD_LABELS[codeIndex]);
                if (!fieldCode) {
                  continue;
                }
                const sourceRow = sourceYearCell.row + 1 + codeIndex;
                const rawValue = getCellRawValue(sourceSheet, sourceRow, sourceYearCell.col);
                const textValue = getCellText(sourceSheet, sourceRow, sourceYearCell.col);
                if (isBlankValue(rawValue) && textValue === "") {
                  continue;
                }
                const value = textValue !== "" ? textValue : rawValue;
                const targetCol = resolveInvestmentTargetColumn(targetSheet, targetHeaderRow, yearStartCol, fieldCode, codeIndex);
                if (targetCol === null) {
                  continue;
                }
                setSheetCellText(targetSheet, writeRow, targetCol, value);
              }
            }
          }

          const sourceCaseSheetName = findSheetNameContainingLabel(sourceWorkbook, "수범 기업 #1");
          const sourceCaseSheet = sourceCaseSheetName ? sourceWorkbook.Sheets[sourceCaseSheetName] : undefined;
          if (sourceCaseSheet) {
            const parsedCases = parseBestPracticeCases(sourceCaseSheet);
            const firstContent = parsedCases[0]?.content?.trim();
            if (firstContent) {
              for (let index = 0; index < 3; index += 1) {
                const markerLabel = `수범 기업 #${index + 1}`;
                const markerCell = findCellByLabel(targetSheet, markerLabel);
                if (!markerCell) {
                  continue;
                }
                const caseData = parsedCases[index];
                if (caseData?.company) {
                  setSheetCellText(targetSheet, writeRow, markerCell.col, caseData.company);
                }
                if (caseData?.content) {
                  setSheetCellText(targetSheet, writeRow, markerCell.col + 1, caseData.content);
                }
              }
            }
          }
        }

        if (targetInvesteeSheet) {
          applyInvesteeFirmsToTemplate(sourceWorkbook, targetInvesteeSheet);
        }
      }

      setUnmatchedFiles(unmatched);
      setWorkbook(workingWorkbook);
      setSheetNames(workingWorkbook.SheetNames);
      selectSheet(workingWorkbook, FUND_SHEET_NAME);
    } catch (error) {
      console.error(error);
      setErrorMessage("업로드 파일 반영 중 오류가 발생했습니다.");
    } finally {
      setIsApplying(false);
    }
  };

  useEffect(() => {
    void loadTemplateWorkbook();
  }, []);

  return (
    <main className="container containerWide">
      <section className="analysisSection">
        <h2 className="analysisTitle">{TEMPLATE_FILE_NAME}</h2>

        <div className="uploadRow">
          <input type="file" accept=".xlsx,.xls" multiple onChange={onUploadFileChange} />
          <button type="button" disabled={uploadedFiles.length === 0 || !templateWorkbook || isApplying} onClick={applyManagerInfoToTemplate}>
            {isApplying ? "반영 중..." : "등록"}
          </button>
          <button type="button" disabled={!workbook || isLoading || isApplying || isDownloading} onClick={() => void downloadWorkbook()}>
            다운로드
          </button>
          <button type="button" disabled={isLoading} onClick={() => void loadTemplateWorkbook()}>
            템플릿 새로고침
          </button>
          <span className="uploadHint">
            {uploadedFiles.length > 0 ? `${uploadedFiles.length}개 파일 선택됨` : "업로드 파일을 선택하세요."}
          </span>
        </div>

        {sheetNames.length > 0 && (
          <div className="sheetTabRow">
            {sheetNames.map((name) => (
              <button
                key={name}
                type="button"
                className={`sheetTab ${selectedSheetName === name ? "sheetTabActive" : "sheetTabInactive"}`}
                onClick={() => {
                  if (workbook) {
                    selectSheet(workbook, name);
                  }
                }}
              >
                {name}
              </button>
            ))}
          </div>
        )}

        {unmatchedFiles.length > 0 && (
          <>
            <button type="button" className="unmatchedToggle" onClick={() => setShowUnmatched((v) => !v)}>
              처리 제외된 파일 ({unmatchedFiles.length}개) {showUnmatched ? "▲" : "▼"}
            </button>
            {showUnmatched && (
              <ul className="unmatchedList">
                {unmatchedFiles.map((name) => (
                  <li key={name}>{name}</li>
                ))}
              </ul>
            )}
          </>
        )}

        {isLoading && <p className="analysisMessage">불러오는 중...</p>}
        {!isLoading && errorMessage && <p className="analysisMessage analysisMessageError">{errorMessage}</p>}

        {!isLoading && !errorMessage && sheetHtml && (
          <div className="originalPreviewViewport">
            <iframe title={`엑셀 미리보기-${selectedSheetName}`} className="originalPreviewFrame" srcDoc={sheetHtml} />
          </div>
        )}
      </section>
    </main>
  );
}
