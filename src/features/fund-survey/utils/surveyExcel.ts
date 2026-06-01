import * as XLSX from "xlsx";
import { TEXT } from "../constants";
import type { FundRow } from "../types";

export function normalize(value: string): string {
  return value.replace(/\s+/g, "");
}

export function normalizeMatch(value: string): string {
  return value.toLowerCase().replace(/[^0-9a-z\uac00-\ud7a3]/g, "");
}

export function normalizeCompanyMatch(value: string): string {
  const withoutParenTokens = value
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\([^)]*\)/g, "");

  const withoutCorpTokens = withoutParenTokens
    .replace(/\(\s*[주유]\s*\)/g, "")
    .replace(/㈜/g, "")
    .replace(/주식회사/g, "")
    .replace(/유한회사/g, "")
    .replace(/유한책임회사/g, "")
    .replace(/합자회사/g, "")
    .replace(/합명회사/g, "");

  const normalized = normalizeMatch(withoutCorpTokens);

  // ?? ? ?? ???? ??? ?? ???? ?? ?? ??
  if (
    normalized.includes("엘앤에스벤처캐피탈") ||
    normalized.includes("엘엔에스벤처캐피탈") ||
    normalized.includes("앨엔에스벤처캐피탈") ||
    normalized.includes("ls벤처캐피탈") ||
    normalized.includes("lns벤처캐피탈")
  ) {
    return "alias_lns_vc";
  }

  if (
    normalized.includes("아이엠투자파트너스") ||
    normalized.includes("하이투자파트너스")
  ) {
    return "alias_im_vc";
  }

  if (
    normalized === "buh" ||
    normalized.includes("비유에이치") ||
    normalized.includes("비유에이취")
  ) {
    return "alias_buh";
  }

  return normalized;
}

export function normalizeFundMatch(value: string): string {
  const normalizedSource = value
    .normalize("NFKC")
    .toLowerCase()
    // "1호" vs "제1호" 같은 순번 표기 차이를 통일
    .replace(/제\s*(?=\d)/g, "");

  return normalizeMatch(normalizedSource);
}

export function makeMatchKey(company: string, fund: string): string {
  return `${normalizeCompanyMatch(company)}|${normalizeFundMatch(fund)}`;
}

export function findColumn(headers: string[], target: string): string | undefined {
  const normalizedTarget = normalize(target);
  return (
    headers.find((header) => normalize(header) === normalizedTarget) ||
    headers.find((header) => normalize(header).includes(normalizedTarget))
  );
}

export function getCellText(sheet: XLSX.WorkSheet, rowIndex: number, colIndex: number): string {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[address];
  if (!cell) {
    return "";
  }
  return String(cell.w ?? cell.v ?? "").trim();
}

export function normalizeLabel(value: string): string {
  return value.toLowerCase().replace(/[\s.]/g, "");
}

export function getRightValueFromColumnALabel(sheet: XLSX.WorkSheet, targets: string[]): string {
  const ref = sheet["!ref"];
  if (!ref) {
    return "";
  }

  const normalizedTargets = new Set(targets.map((target) => normalizeLabel(target)));
  const range = XLSX.utils.decode_range(ref);
  const rowInfo = sheet["!rows"] ?? [];

  for (let r = range.s.r; r <= range.e.r; r += 1) {
    if (rowInfo[r]?.hidden) {
      continue;
    }

    const label = getCellText(sheet, r, 0);
    if (!label || !normalizedTargets.has(normalizeLabel(label))) {
      continue;
    }

    for (let c = 1; c <= range.e.c; c += 1) {
      const value = getCellText(sheet, r, c);
      if (value) {
        return value;
      }
    }
  }

  return "";
}

export function extractSurveyCompany(sheet: XLSX.WorkSheet): string {
  return getRightValueFromColumnALabel(sheet, ["DQ01. 운용사명", "DQ01 운용사명"]);
}

export function extractSurveyFund(sheet: XLSX.WorkSheet): string {
  return getRightValueFromColumnALabel(sheet, ["DQ01. 조합명", "DQ01 조합명"]);
}

export function parseVisibleFundRows(sheet: XLSX.WorkSheet): FundRow[] {
  const ref = sheet["!ref"];
  if (!ref) {
    return [];
  }

  const range = XLSX.utils.decode_range(ref);
  const headerRowIndex = range.s.r;
  const rowInfo = sheet["!rows"] ?? [];
  const colInfo = sheet["!cols"] ?? [];

  const visibleHeaderCols: Array<{ colIndex: number; label: string }> = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    if (colInfo[c]?.hidden) {
      continue;
    }
    const label = getCellText(sheet, headerRowIndex, c);
    if (!label) {
      continue;
    }
    visibleHeaderCols.push({ colIndex: c, label });
  }

  const headerLabels = visibleHeaderCols.map((item) => item.label);
  const companyHeader = findColumn(headerLabels, TEXT.companyKey);
  const fundHeader = findColumn(headerLabels, TEXT.fundKey);
  if (!companyHeader || !fundHeader) {
    return [];
  }

  const companyColIndex = visibleHeaderCols.find((item) => item.label === companyHeader)?.colIndex;
  const fundColIndex = visibleHeaderCols.find((item) => item.label === fundHeader)?.colIndex;
  if (companyColIndex === undefined || fundColIndex === undefined) {
    return [];
  }

  const rows: FundRow[] = [];
  for (let r = headerRowIndex + 1; r <= range.e.r; r += 1) {
    if (rowInfo[r]?.hidden) {
      continue;
    }
    const company = getCellText(sheet, r, companyColIndex);
    const fund = getCellText(sheet, r, fundColIndex);
    if (!company && !fund) {
      continue;
    }
    rows.push({ company, fund });
  }

  return rows;
}

export function sortFundRowsByCompany(rows: FundRow[]): FundRow[] {
  return [...rows].sort((a, b) => {
    const companyOrder = a.company.localeCompare(b.company, "ko-KR");
    if (companyOrder !== 0) {
      return companyOrder;
    }
    return a.fund.localeCompare(b.fund, "ko-KR");
  });
}

export function countInvesteeCompanies(secondSheet: XLSX.WorkSheet | undefined): number {
  if (!secondSheet || !secondSheet["!ref"]) {
    return 0;
  }

  const dataStartRow = 8;
  const companyColIndex = 5; // F column, DQ01(기업명)
  const range = XLSX.utils.decode_range(secondSheet["!ref"]);
  let count = 0;

  for (let row = dataStartRow - 1; row <= range.e.r; row += 1) {
    const value = getCellText(secondSheet, row, companyColIndex);
    if (value) {
      count += 1;
    }
  }
  return count;
}

export function autoFitColumns(sheet: XLSX.WorkSheet, rows: Array<Record<string, string | number>>) {
  if (!rows.length) {
    return;
  }

  const headers = Object.keys(rows[0]);
  const filterIconPadding = 2.5;
  const cellPadding = 1.2;

  const measureWidth = (text: string): number => {
    let len = 0;
    for (const ch of text) {
      if (/[ -~]/.test(ch)) {
        len += 1;
      } else if (/\s/.test(ch)) {
        len += 0.7;
      } else {
        len += 1.9;
      }
    }
    return len;
  };

  const widths = headers.map((header) => measureWidth(header) + filterIconPadding);
  for (const row of rows) {
    headers.forEach((header, index) => {
      const valueWidth = measureWidth(String(row[header] ?? "")) + cellPadding;
      if (valueWidth > widths[index]) {
        widths[index] = valueWidth;
      }
    });
  }

  sheet["!cols"] = widths.map((width) => ({
    wch: Math.ceil(width + 1)
  }));
}

export function setHeaderAutoFilter(sheet: XLSX.WorkSheet) {
  if (!sheet["!ref"]) {
    return;
  }
  sheet["!autofilter"] = { ref: sheet["!ref"] };
}
