import { useState } from "react";
import FundSurveyPage from "../features/fund-survey/FundSurveyPage";
import PerformanceAnalysisPage from "../features/performance-analysis/PerformanceAnalysisPage";
import "./App.css";

type TabId = "survey" | "analysis";

export default function App() {
  const [activeTab, setActiveTab] = useState<TabId>("survey");

  return (
    <div className="appLayout">
      <div className="tabRow">
        <button
          type="button"
          className={`tabButton ${activeTab === "survey" ? "tabButtonActive" : "tabButtonInactive"}`}
          onClick={() => setActiveTab("survey")}
        >
          설문응답 조회
        </button>
        <button
          type="button"
          className={`tabButton ${activeTab === "analysis" ? "tabButtonActive" : "tabButtonInactive"}`}
          onClick={() => setActiveTab("analysis")}
        >
          성과분석조사
        </button>
      </div>

      {activeTab === "survey" ? <FundSurveyPage /> : <PerformanceAnalysisPage />}
    </div>
  );
}
