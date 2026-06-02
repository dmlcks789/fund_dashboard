import { ChangeEvent, useEffect, useState } from "react";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";
import "./PerformanceAnalysisPage.css";
import {
  TEMPLATE_FILE_NAME,
  TEMPLATE_URL,
  FUND_SHEET_NAME,
  INVESTEE_SHEET_NAME,
  INVESTEE_DATA_START_ROW,
  MANAGEMENT_FIELD_LABELS,
  ASSOCIATION_FIELD_LABELS,
  INVESTMENT_FIELD_LABELS,
  IPO_FIELD_LABELS,
  INVESTMENT_YEARS,
} from "./constants";
import {
  getCellText,
  getCellRawValue,
  isBlankValue,
  labelsMatch,
  findCellByLabel,
  findSheetNameContainingLabel,
  extractFieldCode,
  setSheetCellText,
} from "./utils/excelCellUtils";
import {
  findRowToWrite,
  applyLabelMappedSection,
  findInvestmentTargetLayout,
  findRowByColumnAFieldCode,
  findYearCellNearRow,
  resolveInvestmentTargetColumn,
  parseBestPracticeCases,
} from "./utils/fundSheetUtils";
import { applyInvesteeFirmsToTemplate } from "./utils/investeeSheetUtils";
import {
  buildSheetHtml,
  cloneWorkbook,
  getWorkbookValueChanges,
  coercePrimitiveForExcel,
  extendSheetStyles,
} from "./utils/workbookUtils";

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
  const [applyProgress, setApplyProgress] = useState({ current: 0, total: 0 });
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
      const filename = `성과분석조사_취합본.xlsx`;

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
    setApplyProgress({ current: 0, total: uploadedFiles.length });
    try {
      const workingWorkbook = cloneWorkbook(templateWorkbook);
      const targetSheet = workingWorkbook.Sheets[FUND_SHEET_NAME];
      if (!targetSheet) {
        setErrorMessage("템플릿에서 펀드 조합현황 시트를 찾지 못했습니다.");
        return;
      }
      const targetInvesteeSheet = workingWorkbook.Sheets[INVESTEE_SHEET_NAME];

      const unmatched: string[] = [];

      for (let fileIdx = 0; fileIdx < uploadedFiles.length; fileIdx++) {
        const sourceFile = uploadedFiles[fileIdx];
        setApplyProgress({ current: fileIdx + 1, total: uploadedFiles.length });
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
        <div className="uploadRow">
          <input type="file" accept=".xlsx,.xls" multiple onChange={onUploadFileChange} />
          <button type="button" disabled={uploadedFiles.length === 0 || !templateWorkbook || isApplying} onClick={applyManagerInfoToTemplate}>
            {isApplying ? "반영 중..." : "업로드"}
          </button>
          <button type="button" disabled={!workbook || isLoading || isApplying || isDownloading} onClick={() => void downloadWorkbook()}>
            다운로드
          </button>
          <button type="button" disabled={isLoading} onClick={() => void loadTemplateWorkbook()}>
            새로고침
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

        {isApplying && applyProgress.total > 0 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <span className="progressLabel">
              처리 중... {applyProgress.current} / {applyProgress.total}개
            </span>
            <div className="progressBar">
              <div
                className="progressBarFill"
                style={{ width: `${Math.round((applyProgress.current / applyProgress.total) * 100)}%` }}
              />
            </div>
          </div>
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
