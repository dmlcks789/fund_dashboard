import * as XLSX from "xlsx";
import type { CellPrimitive, InvestmentYear } from "../constants";
import { INVESTMENT_YEARS } from "../constants";
import {
  getCellText,
  labelsMatch,
  normalizeLabel,
  findCellByLabel,
  extractFieldCode,
  findFirstMeaningfulValueRightOfCell,
  findFirstNonEmptyCellInRow,
  findLabelInRowWindow,
  findValueRightOfColumnALabel,
  getCellRawValue,
  findValueInRowRangeByFieldCode,
  findValueInRowRangeByLabel,
  isBlankValue,
  isStrictYearToken,
  setSheetCellText,
} from "./excelCellUtils";

// ── 수범기업 케이스 ──────────────────────────────────────────────

export type ParsedCase = {
  company: string;
  content: string;
};

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

export function parseBestPracticeCases(sourceSheet: XLSX.WorkSheet): ParsedCase[] {
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

// ── 헤더/컬럼 탐색 ──────────────────────────────────────────────

// 템플릿 데이터 시작 행(row 3) 이전의 헤더 구역(row 0~2)을 열 기준으로 스캔.
// 세로 병합 등으로 헤더가 여러 행에 걸쳐 있어도 올바르게 탐색한다.
const HEADER_SEARCH_MAX_ROW = 2; // 0-indexed, row 3부터 데이터

export function findColumnByHeaderLabel(sheet: XLSX.WorkSheet, headerLabel: string): number | null {
  const ref = sheet["!ref"];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);
  const maxRow = Math.min(range.e.r, HEADER_SEARCH_MAX_ROW);

  for (let col = range.s.c; col <= range.e.c; col += 1) {
    for (let row = range.s.r; row <= maxRow; row += 1) {
      const text = getCellText(sheet, row, col);
      if (text && labelsMatch(text, headerLabel)) {
        return col;
      }
    }
  }
  return null;
}

export function findColumnByHeaderCode(sheet: XLSX.WorkSheet, fieldCode: string): number | null {
  const ref = sheet["!ref"];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);
  const maxRow = Math.min(range.e.r, HEADER_SEARCH_MAX_ROW);

  for (let col = range.s.c; col <= range.e.c; col += 1) {
    for (let row = range.s.r; row <= maxRow; row += 1) {
      const text = getCellText(sheet, row, col);
      if (extractFieldCode(text) === fieldCode) {
        return col;
      }
    }
  }
  return null;
}

export function findColumnByHeaderCodeOccurrence(
  sheet: XLSX.WorkSheet,
  fieldCode: string,
  occurrenceIndex: number
): number | null {
  const ref = sheet["!ref"];
  if (!ref || occurrenceIndex < 0) return null;
  const range = XLSX.utils.decode_range(ref);
  const maxRow = Math.min(range.e.r, HEADER_SEARCH_MAX_ROW);

  let matchedIndex = 0;
  // 열을 기준으로 스캔해야 열 순서대로 출현 횟수를 올바르게 셀 수 있다.
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    for (let row = range.s.r; row <= maxRow; row += 1) {
      const text = getCellText(sheet, row, col);
      if (extractFieldCode(text) !== fieldCode) continue;
      if (matchedIndex === occurrenceIndex) return col;
      matchedIndex += 1;
      break; // 한 열에서 한 번만 카운트
    }
  }
  return null;
}

export function findInvestmentColumnByYearAndCode(
  sheet: XLSX.WorkSheet,
  year: InvestmentYear,
  fieldCode: string
): number | null {
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

// ── 행 탐색 ─────────────────────────────────────────────────────

export function findRowToWrite(sheet: XLSX.WorkSheet): number {
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

export function findRowByColumnALabel(sheet: XLSX.WorkSheet, targetLabel: string): number | null {
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

export function findRowByColumnAFieldCode(sheet: XLSX.WorkSheet, fieldCode: string): number | null {
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

// ── 필드코드 유틸 ────────────────────────────────────────────────

export function getFieldCodeNumber(fieldCode: string): number | null {
  const matched = fieldCode.match(/^[A-Z]{1,2}(\d{2})$/);
  if (!matched) {
    return null;
  }
  return Number(matched[1]);
}

export function getAssociationOccurrenceIndex(fieldCode: string): number {
  const fieldNumber = getFieldCodeNumber(fieldCode);
  if (fieldCode.startsWith("DQ") && fieldNumber !== null && fieldNumber <= 11) {
    return 1;
  }
  return 0;
}

// ── 섹션 일괄 적용 ───────────────────────────────────────────────

export function applyLabelMappedSection(
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

// ── 연도 컬럼 탐색 ───────────────────────────────────────────────

export function findYearColumnsNearAnchorRow(
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

export function findYearColumnsInSingleRow(
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

export function findYearColumnsAroundRow(
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

export function resolveInvestmentTargetColumn(
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

export function findYearCellNearRow(
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

export function findInvestmentTargetLayout(
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

export function findCellByFieldCodeInColumnRange(
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

export function buildYearColumnRanges(
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

// ── 셀 쓰기 위치 탐색 ────────────────────────────────────────────

export function isLikelyHorizontalHeaderRow(sheet: XLSX.WorkSheet, row: number, col: number): boolean {
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

export function findWriteCellRightOfLabel(
  sheet: XLSX.WorkSheet,
  labelRow: number,
  labelCol: number
): { row: number; col: number } {
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

export function writeValueAtLabelTarget(
  sheet: XLSX.WorkSheet,
  labelCell: { row: number; col: number },
  value: CellPrimitive
) {
  if (isLikelyHorizontalHeaderRow(sheet, labelCell.row, labelCell.col)) {
    setSheetCellText(sheet, labelCell.row + 1, labelCell.col, value);
    return;
  }
  const writeCell = findWriteCellRightOfLabel(sheet, labelCell.row, labelCell.col);
  setSheetCellText(sheet, writeCell.row, writeCell.col, value);
}

// ── 운용사 정보 소스 파싱 ────────────────────────────────────────────

/**
 * 소스 시트에서 "DQ01. 운용사명" 셀의 행/열 위치를 모두 반환.
 * Co-GP이면 2개, 일반이면 1개.
 */
function findManagementDQ01Positions(
  sheet: XLSX.WorkSheet
): Array<{ row: number; labelCol: number }> {
  const ref = sheet["!ref"];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);
  const positions: Array<{ row: number; labelCol: number }> = [];

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      // DQ01 + "운용사명" → 관리정보 DQ01 (조합명과 구분)
      if (extractFieldCode(text) === "DQ01" && labelsMatch(text, "운용사명")) {
        positions.push({ row, labelCol: col });
        break;
      }
    }
  }
  return positions;
}

/**
 * 소스 시트에서 '운용사 정보' 섹션을 모두 찾아 각 섹션의 fieldLabels 값을 반환.
 * 전략: "DQ01. 운용사명" 행을 기준으로 11행을 위치 기반으로 읽는다.
 * (섹션 헤더나 레이블 텍스트 매칭에 의존하지 않아 인코딩 차이에 강함)
 */
export function parseAllManagementSections(
  sheet: XLSX.WorkSheet,
  fieldLabels: readonly string[]
): Array<Record<string, CellPrimitive>> {
  const dq01Positions = findManagementDQ01Positions(sheet);
  if (dq01Positions.length === 0) return [];

  const ref = sheet["!ref"];
  const range = ref ? XLSX.utils.decode_range(ref) : null;

  return dq01Positions.map(({ row: startRow, labelCol }) => {
    const data: Record<string, CellPrimitive> = {};
    for (let i = 0; i < fieldLabels.length && i < 11; i += 1) {
      const dataRow = startRow + i;
      let value: CellPrimitive = null;
      if (range) {
        for (let valueCol = labelCol + 1; valueCol <= range.e.c; valueCol += 1) {
          const rawValue = getCellRawValue(sheet, dataRow, valueCol);
          const textValue = getCellText(sheet, dataRow, valueCol);
          if (!isBlankValue(rawValue) || textValue !== "") {
            value = textValue !== "" ? textValue : rawValue;
            break;
          }
        }
      }
      data[fieldLabels[i]] = value;
    }
    return data;
  });
}

// ── 타겟 템플릿 관리정보 컬럼 탐색 ──────────────────────────────────

/**
 * 타겟 템플릿에서 "DQ01. 운용사명"의 n번째 출현 컬럼을 기준으로
 * DQ01~DQ11 각각의 컬럼 번호를 반환.
 * "1. 운용사 정보"(occurrence=0), "1-1. 운용사 정보"(occurrence=1)
 */
export function findManagementTargetCols(
  sheet: XLSX.WorkSheet,
  sectionOccurrence: number
): Record<string, number> {
  const ref = sheet["!ref"];
  if (!ref) return {};
  const range = XLSX.utils.decode_range(ref);
  const maxRow = Math.min(range.e.r, HEADER_SEARCH_MAX_ROW + 1);

  // "DQ01. 운용사명"이 있는 컬럼 목록 (left→right 순)
  const dq01Cols: number[] = [];
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    for (let row = range.s.r; row <= maxRow; row += 1) {
      const text = getCellText(sheet, row, col);
      if (extractFieldCode(text) === "DQ01" && labelsMatch(text, "운용사명")) {
        dq01Cols.push(col);
        break;
      }
    }
  }

  const startCol = dq01Cols[sectionOccurrence];
  if (startCol === undefined) return {};

  // 다음 "DQ01. 운용사명" 직전까지가 이 섹션의 컬럼 범위
  const endCol = dq01Cols[sectionOccurrence + 1] !== undefined
    ? dq01Cols[sectionOccurrence + 1] - 1
    : range.e.c;

  const result: Record<string, number> = {};
  for (let i = 1; i <= 11; i += 1) {
    const code = `DQ${String(i).padStart(2, "0")}`;
    for (let col = startCol; col <= Math.min(endCol, range.e.c); col += 1) {
      let found = false;
      for (let row = range.s.r; row <= maxRow; row += 1) {
        if (extractFieldCode(getCellText(sheet, row, col)) === code) {
          result[code] = col;
          found = true;
          break;
        }
      }
      if (found) break;
    }
  }
  return result;
}

// ── 조합 정보 파싱 ──────────────────────────────────────────────────

function findAssociationDQ01Position(
  sheet: XLSX.WorkSheet
): { row: number; labelCol: number } | null {
  const ref = sheet["!ref"];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (extractFieldCode(text) === "DQ01" && labelsMatch(text, "조합명")) {
        return { row, labelCol: col };
      }
    }
  }
  return null;
}

/**
 * 소스 시트에서 조합 정보 DQ01~DQ13 값을 위치 기반으로 읽어 반환.
 */
export function parseAssociationSection(
  sheet: XLSX.WorkSheet,
  fieldLabels: readonly string[]
): Record<string, CellPrimitive> {
  const position = findAssociationDQ01Position(sheet);
  if (!position) return {};

  const { row: startRow, labelCol } = position;
  const ref = sheet["!ref"];
  const range = ref ? XLSX.utils.decode_range(ref) : null;
  const data: Record<string, CellPrimitive> = {};

  for (let i = 0; i < fieldLabels.length && i < 13; i += 1) {
    const dataRow = startRow + i;
    let value: CellPrimitive = null;
    if (range) {
      for (let valueCol = labelCol + 1; valueCol <= range.e.c; valueCol += 1) {
        const rawValue = getCellRawValue(sheet, dataRow, valueCol);
        const textValue = getCellText(sheet, dataRow, valueCol);
        if (!isBlankValue(rawValue) || textValue !== "") {
          value = textValue !== "" ? textValue : rawValue;
          break;
        }
      }
    }
    data[fieldLabels[i]] = value;
  }
  return data;
}

/**
 * 타겟 템플릿에서 조합 정보 DQ01~DQ13 컬럼 번호를 반환.
 */
export function findAssociationTargetCols(sheet: XLSX.WorkSheet): Record<string, number> {
  const ref = sheet["!ref"];
  if (!ref) return {};
  const range = XLSX.utils.decode_range(ref);
  const maxRow = Math.min(range.e.r, HEADER_SEARCH_MAX_ROW + 1);

  let startCol: number | undefined;
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    for (let row = range.s.r; row <= maxRow; row += 1) {
      const text = getCellText(sheet, row, col);
      if (extractFieldCode(text) === "DQ01" && labelsMatch(text, "조합명")) {
        startCol = col;
        break;
      }
    }
    if (startCol !== undefined) break;
  }
  if (startCol === undefined) return {};

  const result: Record<string, number> = {};
  for (let i = 1; i <= 13; i += 1) {
    const code = `DQ${String(i).padStart(2, "0")}`;
    for (let col = startCol; col <= range.e.c; col += 1) {
      let found = false;
      for (let row = range.s.r; row <= maxRow; row += 1) {
        if (extractFieldCode(getCellText(sheet, row, col)) === code) {
          result[code] = col;
          found = true;
          break;
        }
      }
      if (found) break;
    }
  }
  return result;
}

// ── 투자 현황 파싱 ──────────────────────────────────────────────────

function isYearHeader(text: string, year: string): boolean {
  const compact = text.replace(/\s/g, "");
  return compact === year || compact === `${year}년`;
}

function findInvestmentSectionHeaderRow(sheet: XLSX.WorkSheet): {
  headerRow: number;
  yearCols: Partial<Record<InvestmentYear, number>>;
} | null {
  const ref = sheet["!ref"];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    let hasGubun = false;
    const yearCols: Partial<Record<InvestmentYear, number>> = {};

    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text) continue;
      if (normalizeLabel(text) === "구분") hasGubun = true;
      for (const year of INVESTMENT_YEARS) {
        if (isYearHeader(text, year)) yearCols[year] = col;
      }
    }

    if (hasGubun && Object.keys(yearCols).length >= 1) {
      return { headerRow: row, yearCols };
    }
  }
  return null;
}

/**
 * 소스 시트에서 투자 현황(A01~A14) 값을 연도 컬럼 기준으로 읽어 반환.
 * { "2025": [A01값, A02값, ..., A14값], "2026": [...], "2027": [...] }
 */
export function parseInvestmentData(
  sheet: XLSX.WorkSheet
): Partial<Record<InvestmentYear, CellPrimitive[]>> {
  const layout = findInvestmentSectionHeaderRow(sheet);
  if (!layout) return {};

  const { headerRow, yearCols } = layout;
  const result: Partial<Record<InvestmentYear, CellPrimitive[]>> = {};

  for (const year of INVESTMENT_YEARS) {
    const col = yearCols[year];
    if (col === undefined) continue;

    const values: CellPrimitive[] = [];
    for (let i = 1; i <= 14; i += 1) {
      const dataRow = headerRow + i;
      const rawValue = getCellRawValue(sheet, dataRow, col);
      const textValue = getCellText(sheet, dataRow, col);
      let value: CellPrimitive = null;
      if (!isBlankValue(rawValue) || textValue !== "") {
        value = textValue !== "" ? textValue : rawValue;
      }
      values.push(value);
    }
    result[year] = values;
  }

  return result;
}

// ── IPO 현황 파싱 ────────────────────────────────────────────────────

/**
 * 소스 시트에서 "A15." 행을 찾아 A15~A17 값을 위치 기반으로 읽어 반환.
 */
export function parseIPOSection(
  sheet: XLSX.WorkSheet,
  fieldLabels: readonly string[]
): Record<string, CellPrimitive> {
  const ref = sheet["!ref"];
  if (!ref) return {};
  const range = XLSX.utils.decode_range(ref);

  let startRow: number | undefined;
  let labelCol: number | undefined;

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      if (extractFieldCode(getCellText(sheet, row, col)) === "A15") {
        startRow = row;
        labelCol = col;
        break;
      }
    }
    if (startRow !== undefined) break;
  }

  if (startRow === undefined || labelCol === undefined) return {};

  const data: Record<string, CellPrimitive> = {};
  for (let i = 0; i < fieldLabels.length && i < 3; i += 1) {
    const dataRow = startRow + i;
    let value: CellPrimitive = null;
    for (let valueCol = labelCol + 1; valueCol <= range.e.c; valueCol += 1) {
      const rawValue = getCellRawValue(sheet, dataRow, valueCol);
      const textValue = getCellText(sheet, dataRow, valueCol);
      if (!isBlankValue(rawValue) || textValue !== "") {
        value = textValue !== "" ? textValue : rawValue;
        break;
      }
    }
    data[fieldLabels[i]] = value;
  }
  return data;
}

/**
 * 타겟 템플릿에서 A15~A17 컬럼 번호를 반환.
 */
export function findIPOTargetCols(sheet: XLSX.WorkSheet): Record<string, number> {
  const ref = sheet["!ref"];
  if (!ref) return {};
  const range = XLSX.utils.decode_range(ref);
  const maxRow = Math.min(range.e.r, HEADER_SEARCH_MAX_ROW + 1);

  let startCol: number | undefined;
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    for (let row = range.s.r; row <= maxRow; row += 1) {
      if (extractFieldCode(getCellText(sheet, row, col)) === "A15") {
        startCol = col;
        break;
      }
    }
    if (startCol !== undefined) break;
  }
  if (startCol === undefined) return {};

  const result: Record<string, number> = {};
  for (const code of ["A15", "A16", "A17"]) {
    for (let col = startCol; col <= range.e.c; col += 1) {
      let found = false;
      for (let row = range.s.r; row <= maxRow; row += 1) {
        if (extractFieldCode(getCellText(sheet, row, col)) === code) {
          result[code] = col;
          found = true;
          break;
        }
      }
      if (found) break;
    }
  }
  return result;
}
