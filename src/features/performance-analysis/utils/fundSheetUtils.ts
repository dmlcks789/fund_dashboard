import * as XLSX from "xlsx";
import type { CellPrimitive, InvestmentYear } from "../constants";
import { INVESTMENT_YEARS } from "../constants";
import {
  getCellText,
  labelsMatch,
  findCellByLabel,
  extractFieldCode,
  findFirstMeaningfulValueRightOfCell,
  findFirstNonEmptyCellInRow,
  findLabelInRowWindow,
  findValueRightOfColumnALabel,
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

export function findColumnByHeaderLabel(sheet: XLSX.WorkSheet, headerLabel: string): number | null {
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

export function findColumnByHeaderCode(sheet: XLSX.WorkSheet, fieldCode: string): number | null {
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

export function findColumnByHeaderCodeOccurrence(
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
