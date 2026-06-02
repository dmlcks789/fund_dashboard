import ExcelJS from "exceljs";
import * as XLSX from "xlsx";
import type { CellPrimitive } from "../constants";
import { INVESTEE_DATA_START_ROW } from "../constants";
import { getCellText, getCellRawValue, isBlankValue } from "./excelCellUtils";

export function buildSheetHtml(sheetName: string, sheet: XLSX.WorkSheet): string {
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

export function cloneWorkbook(workbook: XLSX.WorkBook): XLSX.WorkBook {
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

export function getWorkbookValueChanges(current: XLSX.WorkBook, base: XLSX.WorkBook) {
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

export function coercePrimitiveForExcel(value: CellPrimitive): string | number | boolean | Date | null {
  if (value === null || typeof value === "number" || typeof value === "boolean" || value instanceof Date) {
    return value;
  }
  const trimmed = value.trim();
  if (trimmed === "") {
    return null;
  }
  return trimmed;
}

export function extendSheetStyles(worksheet: ExcelJS.Worksheet, firstDataRow: number): void {
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
    const firstRowDvs: Array<{ col: string; dv: unknown }> = [];
    for (const [ref, dv] of Object.entries(model)) {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if (m && parseInt(m[2]) === firstDataRow) {
        firstRowDvs.push({ col: m[1], dv });
      }
    }

    const existingRows = new Set<number>();
    for (const ref of Object.keys(model)) {
      const m = ref.match(/^[A-Z]+(\d+)$/);
      if (m) existingRows.add(parseInt(m[1]));
    }

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
