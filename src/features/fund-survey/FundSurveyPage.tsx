import { ChangeEvent, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import FundTable from "./components/FundTable";
import UnmatchedSurveyModal from "./components/UnmatchedSurveyModal";
import { EXPORT_FILE_NAME, TEXT } from "./constants";
import type { FundRow, Status } from "./types";
import {
  autoFitColumns,
  countInvesteeCompanies,
  extractSurveyCompany,
  extractSurveyFund,
  makeMatchKey,
  normalizeCompanyMatch,
  normalizeFundMatch,
  parseVisibleFundRows,
  setHeaderAutoFilter,
  sortFundRowsByCompany
} from "./utils/surveyExcel";
import "./FundSurveyPage.css";

export default function FundSurveyPage() {
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

      <FundTable
        rows={rows}
        surveyMatchKeys={surveyMatchKeys}
        surveyFundInvesteeMap={surveyFundInvesteeMap}
        investeeCountMap={investeeCountMap}
      />
      <UnmatchedSurveyModal
        isOpen={showUnmatchedModal}
        fileNames={unmatchedSurveyFiles}
        onClose={() => setShowUnmatchedModal(false)}
      />
    </main>
  );
}

