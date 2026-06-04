import * as XLSX from "xlsx";
import type { CellPrimitive, InvestmentYear } from "../constants";

export function normalizeLabel(value: string): string {
  return value.toLowerCase().replace(/[^0-9a-z가-힣]/g, "");
}

export function getCellText(sheet: XLSX.WorkSheet, row: number, col: number): string {
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

export function getCellRawValue(sheet: XLSX.WorkSheet, row: number, col: number): CellPrimitive {
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

export function formatNumberForDisplay(value: number): string {
  return new Intl.NumberFormat("en-US", {
    useGrouping: false,
    maximumFractionDigits: 20
  }).format(value);
}

export function isBlankValue(value: CellPrimitive): boolean {
  return value === null || (typeof value === "string" && value.trim() === "");
}

export function labelsMatch(a: string, b: string): boolean {
  const na = normalizeLabel(a);
  const nb = normalizeLabel(b);
  return na === nb || na.includes(nb) || nb.includes(na);
}

export function findCellByLabel(
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

export function findSheetNameContainingLabel(workbook: XLSX.WorkBook, targetLabel: string): string | null {
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

export function findSheetNameByPreferredOrder(workbook: XLSX.WorkBook, fundSheetName: string): string | null {
  return workbook.SheetNames.find((name) => labelsMatch(name, fundSheetName)) ?? workbook.SheetNames[0] ?? null;
}

export function extractFieldCode(label: string): string {
  const normalized = label.trim().toUpperCase();
  const matched = normalized.match(/^([A-Z]{1,2}\d{2})\s*\./);
  return matched?.[1] ?? "";
}

export function findFirstValueRightOfCell(
  sheet: XLSX.WorkSheet,
  row: number,
  col: number
): { col: number; value: string } | null {
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

export function findFirstMeaningfulValueRightOfCell(
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

export function findFirstNonEmptyCellInRow(
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

export function findLabelInRowWindow(
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

export function findAllValuesRightOfColumnALabel(sheet: XLSX.WorkSheet, targetLabel: string): string {
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

export function findValueRightOfColumnALabel(sheet: XLSX.WorkSheet, targetLabel: string): CellPrimitive {
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

/**
 * 지정 행 범위 내에서 DQ 필드코드(예: "DQ08")로 레이블 셀을 찾아 우측 첫 값을 반환.
 * 한글 텍스트 매칭보다 견고하다.
 */
export function findValueInRowRangeByFieldCode(
  sheet: XLSX.WorkSheet,
  startRow: number,
  endRow: number,
  fieldCode: string
): CellPrimitive {
  const ref = sheet["!ref"];
  if (!ref || !fieldCode) return null;
  const range = XLSX.utils.decode_range(ref);
  const fromRow = Math.max(range.s.r, startRow);
  const toRow = Math.min(range.e.r, endRow);

  for (let row = fromRow; row <= toRow; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text || extractFieldCode(text) !== fieldCode) continue;
      for (let valueCol = col + 1; valueCol <= range.e.c; valueCol += 1) {
        const rawValue = getCellRawValue(sheet, row, valueCol);
        const textValue = getCellText(sheet, row, valueCol);
        if (!isBlankValue(rawValue) || textValue !== "") {
          return textValue !== "" ? textValue : rawValue;
        }
      }
      return null;
    }
  }
  return null;
}

export function findValueInRowRangeByLabel(
  sheet: XLSX.WorkSheet,
  startRow: number,
  endRow: number,
  targetLabel: string
): CellPrimitive {
  const ref = sheet["!ref"];
  if (!ref) return null;
  const range = XLSX.utils.decode_range(ref);
  const fromRow = Math.max(range.s.r, startRow);
  const toRow = Math.min(range.e.r, endRow);

  for (let row = fromRow; row <= toRow; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const text = getCellText(sheet, row, col);
      if (!text || !labelsMatch(text, targetLabel)) continue;
      for (let valueCol = col + 1; valueCol <= range.e.c; valueCol += 1) {
        const rawValue = getCellRawValue(sheet, row, valueCol);
        const textValue = getCellText(sheet, row, valueCol);
        if (!isBlankValue(rawValue) || textValue !== "") {
          return textValue !== "" ? textValue : rawValue;
        }
      }
      return null;
    }
  }
  return null;
}

export function findValueRightOfColumnAFieldCodeOccurrence(
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

export function isStrictYearToken(text: string, year: InvestmentYear): boolean {
  const compact = text.replace(/\s/g, "");
  const matchedYear = compact.match(/(2025|2026|2027)/)?.[1];
  return matchedYear === year;
}

export function setSheetCellText(sheet: XLSX.WorkSheet, row: number, col: number, value: CellPrimitive) {
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
  if (row < range.s.r) range.s.r = row;
  if (col < range.s.c) range.s.c = col;
  if (row > range.e.r) range.e.r = row;
  if (col > range.e.c) range.e.c = col;
  sheet["!ref"] = XLSX.utils.encode_range(range);
}
