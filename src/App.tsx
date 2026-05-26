import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

type FundRow = {
  company: string;
  fund: string;
};

type StatusTone = "success" | "info" | "";
type Status = { message: string; tone: StatusTone };

const TEXT = {
  listLabel: "\uB9AC\uC2A4\uD2B8:",
  surveyLabel: "\uC124\uBB38\uC751\uB2F5:",
  load: "\uBD88\uB7EC\uC624\uAE30",
  loading: "\uBD88\uB7EC\uC624\uB294 \uC911...",
  surveyLoading: "\uCC98\uB9AC \uC911...",
  companyHeader: "\uC6B4\uC6A9\uC0AC",
  fundHeader: "\uD380\uB4DC\uBA85",
  surveyHeader: "\uC124\uBB38\uC751\uB2F5 \uC5EC\uBD80",
  investeeHeader: "\uD53C\uD22C\uC790\uAE30\uC5C5 \uC218",
  surveyDone: "\uC751\uB2F5 \uC644\uB8CC",
  surveyPending: "",
  download: "\uC5D1\uC140 \uB2E4\uC6B4 \uBC1B\uAE30",
  noHeader: "No",
  companyKey: "\uC6B4\uC6A9\uC0AC",
  fundKey: "\uD380\uB4DC\uBA85",
  pickFileFirst: "\uD30C\uC77C \uC120\uD0DD \uD6C4 \uBD88\uB7EC\uC624\uC138\uC694.",
  loadListFirst: "\uB9AC\uC2A4\uD2B8\uB97C \uBA3C\uC800 \uBD88\uB7EC\uC624\uC138\uC694.",
  sheetNotFound: "\uC2DC\uD2B8\uB97C \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.",
  uploadFailed: "\uC5C5\uB85C\uB4DC \uC2E4\uD328",
  uploadDone: "\uC5C5\uB85C\uB4DC \uC644\uB8CC",
  downloadNoData: "\uB2E4\uC6B4\uB85C\uB4DC\uD560 \uD56D\uBAA9\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.",
  downloadDone: "\uB2E4\uC6B4\uB85C\uB4DC \uC644\uB8CC",
  unmatchedButton: "\uBBF8\uB9E4\uCE6D \uD30C\uC77C",
  unmatchedTitle: "\uB9E4\uCE6D \uC548 \uB41C \uD30C\uC77C \uBAA9\uB85D",
  unmatchedEmpty: "\uBAA8\uB4E0 \uD30C\uC77C\uC774 \uB9E4\uCE6D\uB418\uC5C8\uC2B5\uB2C8\uB2E4.",
  close: "\uB2EB\uAE30",
  overviewCompany: "\uC6B4\uC6A9\uC0AC",
  overviewFund: "\uD380\uB4DC",
  overviewSurvey: "\uC124\uBB38 \uC751\uB2F5",
  overviewRate: "\uC751\uB2F5\uB960",
  countUnit: "\uAC1C",
  overviewSeparator: "\u00B7"
};

const EXPORT_FILE_NAME =
  "\u0032\u0030\u0032\u0036\uB144 \uBD80\uC0B0\uAD11\uC5ED\uC2DC \uD380\uB4DC \uC870\uC131 \uBC0F \uC131\uACFC\uBD84\uC11D_\uC124\uBB38\uC751\uB2F5 \uB9AC\uC2A4\uD2B8.xlsx";

function normalize(value: string): string {
  return value.replace(/\s+/g, "");
}

function normalizeMatch(value: string): string {
  return value.toLowerCase().replace(/[^0-9a-z\uac00-\ud7a3]/g, "");
}

function normalizeCompanyMatch(value: string): string {
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

function normalizeFundMatch(value: string): string {
  const normalizedSource = value
    .normalize("NFKC")
    .toLowerCase()
    // "1호" vs "제1호" 같은 순번 표기 차이를 통일
    .replace(/제\s*(?=\d)/g, "");

  return normalizeMatch(normalizedSource);
}

function makeMatchKey(company: string, fund: string): string {
  return `${normalizeCompanyMatch(company)}|${normalizeFundMatch(fund)}`;
}

function findColumn(headers: string[], target: string): string | undefined {
  const normalizedTarget = normalize(target);
  return (
    headers.find((header) => normalize(header) === normalizedTarget) ||
    headers.find((header) => normalize(header).includes(normalizedTarget))
  );
}

function getCellText(sheet: XLSX.WorkSheet, rowIndex: number, colIndex: number): string {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[address];
  if (!cell) {
    return "";
  }
  return String(cell.w ?? cell.v ?? "").trim();
}

function normalizeLabel(value: string): string {
  return value.toLowerCase().replace(/[\s.]/g, "");
}

function getRightValueFromColumnALabel(sheet: XLSX.WorkSheet, targets: string[]): string {
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

function extractSurveyCompany(sheet: XLSX.WorkSheet): string {
  return getRightValueFromColumnALabel(sheet, ["DQ01. 운용사명", "DQ01 운용사명"]);
}

function extractSurveyFund(sheet: XLSX.WorkSheet): string {
  return getRightValueFromColumnALabel(sheet, ["DQ01. 조합명", "DQ01 조합명"]);
}

function parseVisibleFundRows(sheet: XLSX.WorkSheet): FundRow[] {
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

function sortFundRowsByCompany(rows: FundRow[]): FundRow[] {
  return [...rows].sort((a, b) => {
    const companyOrder = a.company.localeCompare(b.company, "ko-KR");
    if (companyOrder !== 0) {
      return companyOrder;
    }
    return a.fund.localeCompare(b.fund, "ko-KR");
  });
}

function countInvesteeCompanies(secondSheet: XLSX.WorkSheet | undefined): number {
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

function autoFitColumns(sheet: XLSX.WorkSheet, rows: Array<Record<string, string | number>>) {
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

function setHeaderAutoFilter(sheet: XLSX.WorkSheet) {
  if (!sheet["!ref"]) {
    return;
  }
  sheet["!autofilter"] = { ref: sheet["!ref"] };
}

function OverflowTooltipText({ value }: { value: string }) {
  const textRef = useRef<HTMLSpanElement>(null);
  const [isOverflow, setIsOverflow] = useState(false);

  useEffect(() => {
    const element = textRef.current;
    if (!element) {
      return;
    }

    const checkOverflow = () => {
      setIsOverflow(element.scrollWidth > element.clientWidth + 1);
    };

    checkOverflow();
    window.addEventListener("resize", checkOverflow);

    let observer: ResizeObserver | null = null;
    if (typeof ResizeObserver !== "undefined") {
      observer = new ResizeObserver(checkOverflow);
      observer.observe(element);
    }

    return () => {
      window.removeEventListener("resize", checkOverflow);
      observer?.disconnect();
    };
  }, [value]);

  return (
    <span ref={textRef} className="cellEllipsisText" title={isOverflow ? value : undefined}>
      {value}
    </span>
  );
}

export default function App() {
  const [listFile, setListFile] = useState<File | null>(null);
  const [surveyFiles, setSurveyFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<FundRow[]>([]);
  const [loadingList, setLoadingList] = useState(false);
  const [loadingSurvey, setLoadingSurvey] = useState(false);
  const [surveyMatchKeys, setSurveyMatchKeys] = useState<Set<string>>(new Set());
  const [surveyFundInvesteeMap, setSurveyFundInvesteeMap] = useState<Map<string, number>>(new Map());
  const [investeeCountMap, setInvesteeCountMap] = useState<Map<string, number>>(new Map());
  const [listStatus, setListStatus] = useState<Status>({ message: "", tone: "" });
  const [surveyStatus, setSurveyStatus] = useState<Status>({ message: "", tone: "" });
  const [downloadStatus, setDownloadStatus] = useState<Status>({ message: "", tone: "" });
  const [unmatchedSurveyFiles, setUnmatchedSurveyFiles] = useState<string[]>([]);
  const [showUnmatchedModal, setShowUnmatchedModal] = useState(false);

  const completedRows = useMemo(
    () =>
      rows.filter((row) => {
        const key = makeMatchKey(row.company, row.fund);
        const fundKey = normalizeFundMatch(row.fund);
        return surveyMatchKeys.has(key) || surveyFundInvesteeMap.has(fundKey);
      }),
    [rows, surveyMatchKeys, surveyFundInvesteeMap]
  );

  const totalCount = rows.length;
  const completedCount = completedRows.length;
  const uniqueCompanyCount = useMemo(
    () => new Set(rows.map((row) => normalizeCompanyMatch(row.company)).filter(Boolean)).size,
    [rows]
  );
  const responseRate = totalCount ? ((completedCount / totalCount) * 100).toFixed(1) : "0.0";
  const hasAnyRows = rows.length > 0;

  const onListFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setListFile(event.target.files?.[0] ?? null);
  };

  const onSurveyFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setSurveyFiles(Array.from(event.target.files ?? []));
  };

  const loadListFile = async () => {
    if (!listFile) {
      setListStatus({ message: TEXT.pickFileFirst, tone: "info" });
      return;
    }

    setLoadingList(true);
    try {
      const buffer = await listFile.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellStyles: true });
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        setRows([]);
        setListStatus({ message: TEXT.sheetNotFound, tone: "info" });
        return;
      }

      const sheet = workbook.Sheets[firstSheetName];
      const sortedRows = sortFundRowsByCompany(parseVisibleFundRows(sheet));
      setRows(sortedRows);
      setListStatus({
        message: TEXT.uploadDone,
        tone: "success"
      });
    } catch (error) {
      console.error(error);
      setRows([]);
      setListStatus({ message: TEXT.uploadFailed, tone: "info" });
    } finally {
      setLoadingList(false);
    }
  };

  const loadSurveyFile = async () => {
    if (!hasAnyRows) {
      setSurveyStatus({ message: TEXT.loadListFirst, tone: "info" });
      return;
    }

    if (!surveyFiles.length) {
      setSurveyStatus({ message: TEXT.pickFileFirst, tone: "info" });
      return;
    }

    setLoadingSurvey(true);
    try {
      const nextKeys = new Set<string>();
      const nextInvesteeMap = new Map<string, number>();
      const nextFundInvesteeMap = new Map<string, number>();
      const unmatchedFileNames: string[] = [];
      const listKeySet = new Set(rows.map((row) => makeMatchKey(row.company, row.fund)));
      const listFundSet = new Set(rows.map((row) => normalizeFundMatch(row.fund)));

      for (const surveyFile of surveyFiles) {
        const buffer = await surveyFile.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
          continue;
        }

        const firstSheet = workbook.Sheets[firstSheetName];
        const company = extractSurveyCompany(firstSheet);
        const fund = extractSurveyFund(firstSheet);
        if (!fund) {
          unmatchedFileNames.push(surveyFile.name);
          continue;
        }

        const secondSheetName = workbook.SheetNames[1];
        const secondSheet = secondSheetName ? workbook.Sheets[secondSheetName] : undefined;
        const investeeCount = countInvesteeCompanies(secondSheet);
        const fundKey = normalizeFundMatch(fund);
        const matchedByFund = listFundSet.has(fundKey);

        let matchedByKey = false;
        if (company) {
          const key = makeMatchKey(company, fund);
          if (listKeySet.has(key)) {
            matchedByKey = true;
            nextKeys.add(key);
            nextInvesteeMap.set(key, investeeCount);
          }
        }

        if (!matchedByKey && !matchedByFund) {
          unmatchedFileNames.push(surveyFile.name);
          continue;
        }

        nextFundInvesteeMap.set(fundKey, investeeCount);
      }

      setSurveyMatchKeys((prev) => {
        const next = new Set(prev);
        nextKeys.forEach((key) => next.add(key));
        return next;
      });

      setInvesteeCountMap((prev) => {
        const next = new Map(prev);
        nextInvesteeMap.forEach((value, key) => next.set(key, value));
        return next;
      });

      setSurveyFundInvesteeMap((prev) => {
        const next = new Map(prev);
        nextFundInvesteeMap.forEach((value, key) => next.set(key, value));
        return next;
      });
      setUnmatchedSurveyFiles(Array.from(new Set(unmatchedFileNames)));

      setSurveyStatus({
        message: TEXT.uploadDone,
        tone: "success"
      });
    } catch (error) {
      console.error(error);
      setSurveyStatus({ message: TEXT.uploadFailed, tone: "info" });
    } finally {
      setLoadingSurvey(false);
    }
  };

  const downloadList = () => {
    if (!hasAnyRows) {
      setDownloadStatus({ message: TEXT.downloadNoData, tone: "info" });
      return;
    }

    const exportRows = rows.map((row, index) => {
      const key = makeMatchKey(row.company, row.fund);
      const fundKey = normalizeFundMatch(row.fund);
      const isDone = surveyMatchKeys.has(key) || surveyFundInvesteeMap.has(fundKey);
      const investeeCount = investeeCountMap.get(key) ?? surveyFundInvesteeMap.get(fundKey) ?? "";
      return {
        No: index + 1,
        [TEXT.companyHeader]: row.company,
        [TEXT.fundHeader]: row.fund,
        [TEXT.surveyHeader]: isDone ? TEXT.surveyDone : "",
        [TEXT.investeeHeader]: investeeCount
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    autoFitColumns(worksheet, exportRows);
    setHeaderAutoFilter(worksheet);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "list");
    XLSX.writeFile(workbook, EXPORT_FILE_NAME);
    setDownloadStatus({
      message: TEXT.downloadDone,
      tone: "success"
    });
  };

  return (
    <main className="container">
      <div className="uploadRow">
        <label className="uploadLabel">{TEXT.listLabel}</label>
        <input type="file" accept=".xlsx,.xls" onChange={onListFileChange} />
        <button type="button" disabled={loadingList} onClick={loadListFile}>
          {loadingList ? TEXT.loading : TEXT.load}
        </button>
        <span className={`inlineStatus ${listStatus.tone === "success" ? "inlineStatusSuccess" : "inlineStatusInfo"}`}>
          {listStatus.message || "\u00A0"}
        </span>
      </div>

      <div className="uploadRow">
        <label className="uploadLabel">{TEXT.surveyLabel}</label>
        <input type="file" accept=".xlsx,.xls" multiple onChange={onSurveyFileChange} />
        <button type="button" disabled={loadingSurvey || !hasAnyRows} onClick={loadSurveyFile}>
          {loadingSurvey ? TEXT.surveyLoading : TEXT.load}
        </button>
        <span className={`inlineStatus ${surveyStatus.tone === "success" ? "inlineStatusSuccess" : "inlineStatusInfo"}`}>
          {surveyStatus.message || "\u00A0"}
        </span>
      </div>

      <div className="downloadRow">
        <button
          type="button"
          className={unmatchedSurveyFiles.length > 0 ? "unmatchedButtonActive" : "unmatchedButtonInactive"}
          disabled={unmatchedSurveyFiles.length === 0}
          onClick={() => setShowUnmatchedModal(true)}
        >
          {`${TEXT.unmatchedButton} (${unmatchedSurveyFiles.length})`}
        </button>
        <span
          className={`inlineStatus inlineStatusTight ${
            downloadStatus.tone === "success" ? "inlineStatusSuccess" : "inlineStatusInfo"
          }`}
        >
          {downloadStatus.message || "\u00A0"}
        </span>
        <button type="button" disabled={!hasAnyRows} onClick={downloadList}>
          {TEXT.download}
        </button>
      </div>

      <div className="overviewRow">
        <span className="overviewText">
          {`${TEXT.overviewCompany} ${uniqueCompanyCount.toLocaleString()}${TEXT.countUnit} ${TEXT.overviewSeparator} ${TEXT.overviewFund} ${totalCount.toLocaleString()}${TEXT.countUnit} ${TEXT.overviewSeparator} ${TEXT.overviewSurvey} ${completedCount.toLocaleString()}${TEXT.countUnit} (${TEXT.overviewRate} ${responseRate}%)`}
        </span>
      </div>

      <div className="tableViewport">
        <table>
          <thead>
            <tr>
              <th>{TEXT.noHeader}</th>
              <th>{TEXT.companyHeader}</th>
              <th>{TEXT.fundHeader}</th>
              <th>{TEXT.surveyHeader}</th>
              <th>{TEXT.investeeHeader}</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, index) => {
              const key = makeMatchKey(row.company, row.fund);
              const fundKey = normalizeFundMatch(row.fund);
              const isDone = surveyMatchKeys.has(key) || surveyFundInvesteeMap.has(fundKey);
              const investeeCount = investeeCountMap.get(key) ?? surveyFundInvesteeMap.get(fundKey) ?? "";
              return (
                <tr key={`${row.company}-${row.fund}-${index}`}>
                  <td className="noCell">{index + 1}</td>
                  <td>
                    <OverflowTooltipText value={row.company} />
                  </td>
                  <td>
                    <OverflowTooltipText value={row.fund} />
                  </td>
                  <td className="surveyCell">
                    {isDone ? TEXT.surveyDone : TEXT.surveyPending}
                  </td>
                  <td>{investeeCount}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {showUnmatchedModal && (
        <div className="modalBackdrop" onClick={() => setShowUnmatchedModal(false)}>
          <div className="modalCard" onClick={(event) => event.stopPropagation()}>
            <div className="modalHeader">
              <strong>{TEXT.unmatchedTitle}</strong>
            </div>
            <div className="modalBody">
              {unmatchedSurveyFiles.length === 0 ? (
                <p className="modalEmpty">{TEXT.unmatchedEmpty}</p>
              ) : (
                <ul className="modalList">
                  {unmatchedSurveyFiles.map((name) => (
                    <li key={name}>{name}</li>
                  ))}
                </ul>
              )}
            </div>
            <div className="modalFooter">
              <button type="button" onClick={() => setShowUnmatchedModal(false)}>
                {TEXT.close}
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
