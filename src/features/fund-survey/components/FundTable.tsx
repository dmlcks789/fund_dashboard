import { TEXT } from "../constants";
import type { FundRow } from "../types";
import { makeMatchKey, normalizeFundMatch } from "../utils/surveyExcel";
import OverflowTooltipText from "./OverflowTooltipText";

type FundTableProps = {
  rows: FundRow[];
  surveyMatchKeys: Set<string>;
  surveyFundInvesteeMap: Map<string, number>;
  investeeCountMap: Map<string, number>;
};

export default function FundTable({
  rows,
  surveyMatchKeys,
  surveyFundInvesteeMap,
  investeeCountMap
}: FundTableProps) {
  return (
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
                <td className="surveyCell">{isDone ? TEXT.surveyDone : TEXT.surveyPending}</td>
                <td>{investeeCount}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
