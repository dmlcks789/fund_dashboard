import * as XLSX from "xlsx";
import type { CellPrimitive } from "../constants";
import {
  FUND_SHEET_NAME,
  INVESTEE_SHEET_NAME,
  INVESTEE_DATA_START_ROW,
  INVESTEE_SOURCE_COL_OFFSET,
  INVESTEE_SECTION_ROW,
  INVESTEE_SUBHEADER_ROW,
  INVESTEE_FIXED_SRC_COLS,
  TIME_SERIES_SECTION_LABELS,
} from "../constants";
import {
  getCellText,
  getCellRawValue,
  labelsMatch,
  isBlankValue,
  findAllValuesRightOfColumnALabel,
  findValueRightOfColumnALabel,
  findSheetNameContainingLabel,
  setSheetCellText,
} from "./excelCellUtils";

function findSectionSumCol(sheet: XLSX.WorkSheet, sectionStartCol: number): number | null {
  const ref = sheet["!ref"];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);
  for (let col = sectionStartCol + 1; col <= range.e.c; col++) {
    if (labelsMatch(getCellText(sheet, INVESTEE_SUBHEADER_ROW, col), "합계")) return col;
  }
  return null;
}

function findInvesteeSectionStart(sheet: XLSX.WorkSheet, label: string, afterCol: number): number | null {
  const ref = sheet["!ref"];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);
  for (let col = afterCol; col <= range.e.c; col++) {
    const text = getCellText(sheet, INVESTEE_SECTION_ROW, col);
    if (text && labelsMatch(text, label)) return col;
  }
  return null;
}

function buildInvesteeColMap(sourceSheet: XLSX.WorkSheet, targetSheet: XLSX.WorkSheet): Map<number, number> {
  const colMap = new Map<number, number>();

  // 고정 오프셋 구간: 연락처 + DQ코드 (source 1~45 → target 3~47)
  const srcRef = XLSX.utils.decode_range(sourceSheet["!ref"] ?? "A1");
  for (let c = 1; c <= Math.min(INVESTEE_FIXED_SRC_COLS, srcRef.e.c); c++) {
    colMap.set(c, c + INVESTEE_SOURCE_COL_OFFSET);
  }

  // 시계열 섹션: 헤더 텍스트로 섹션 위치 탐색 후 상대 인덱스 매핑
  let srcFrom = INVESTEE_FIXED_SRC_COLS + 1;
  let tgtFrom = INVESTEE_FIXED_SRC_COLS + INVESTEE_SOURCE_COL_OFFSET + 1;
  for (const label of TIME_SERIES_SECTION_LABELS) {
    const srcStart = findInvesteeSectionStart(sourceSheet, label, srcFrom);
    const tgtStart = findInvesteeSectionStart(targetSheet, label, tgtFrom);
    if (srcStart === null || tgtStart === null) continue;
    const srcSum = findSectionSumCol(sourceSheet, srcStart);
    const tgtSum = findSectionSumCol(targetSheet, tgtStart);
    if (srcSum === null || tgtSum === null) continue;

    for (let srcCol = srcStart; srcCol < srcSum; srcCol++) {
      const offset = srcCol - srcStart;
      const tgtCol = tgtStart + offset;
      if (tgtCol < tgtSum) colMap.set(srcCol, tgtCol);
    }
    colMap.set(srcSum, tgtSum);

    srcFrom = srcSum + 1;
    tgtFrom = tgtSum + 1;
  }
  return colMap;
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

export function applyInvesteeFirmsToTemplate(sourceWorkbook: XLSX.WorkBook, targetSheet: XLSX.WorkSheet) {
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

  const colMap = buildInvesteeColMap(sourceSheet, targetSheet);

  const ref = sourceSheet["!ref"];
  if (!ref) return;
  const range = XLSX.utils.decode_range(ref);

  const SRC_COMPANY_COL = 5; // DQ01. 기업명

  for (let srcRow = INVESTEE_DATA_START_ROW; srcRow <= range.e.r; srcRow += 1) {
    if (!getCellText(sourceSheet, srcRow, SRC_COMPANY_COL)) continue;

    const writeRow = findInvesteeWriteRow(targetSheet);
    setSheetCellText(targetSheet, writeRow, 0, writeRow - INVESTEE_DATA_START_ROW + 1);
    if (!isBlankValue(운용사명)) setSheetCellText(targetSheet, writeRow, 1, 운용사명);
    if (!isBlankValue(조합명)) setSheetCellText(targetSheet, writeRow, 2, 조합명);

    for (let srcCol = 1; srcCol <= range.e.c; srcCol += 1) {
      const tgtCol = colMap.get(srcCol);
      if (tgtCol === undefined) continue;
      const rawValue = getCellRawValue(sourceSheet, srcRow, srcCol);
      const textValue = getCellText(sourceSheet, srcRow, srcCol);
      if (isBlankValue(rawValue) && textValue === "") continue;
      setSheetCellText(targetSheet, writeRow, tgtCol, textValue !== "" ? textValue : rawValue);
    }
  }
}
