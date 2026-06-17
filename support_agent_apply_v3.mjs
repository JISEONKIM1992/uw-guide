import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const guidelinePath = String.raw`C:\Users\user\Desktop\신계약가이드라인_공통코어_샘플결과_v3.xlsx`;
const rdPath = String.raw`C:\Users\user\Desktop\RD판.xlsx`;
const outputDir = String.raw`C:\Users\user\Desktop\uw guide\outputs`;
const outputPath = path.join(outputDir, "RD판_지원Agent_v3_룰포함_자동수정.xlsx");

await fs.mkdir(outputDir, { recursive: true });

const fills = {
  normal: "#FFF59D",
  rule: "#F9D5A7",
  info: "#D9EAF7",
  title: "#1F4E78",
  border: "#B8C4D4",
  light: "#FFFDF2",
};

const guidelineBlob = await FileBlob.load(guidelinePath);
const guidelineWb = await SpreadsheetFile.importXlsx(guidelineBlob);
const ruleSheet = guidelineWb.worksheets.getItem("Rule Master");
const ruleValues = ruleSheet.getUsedRange().values;
const ruleHeaders = ruleValues[1];

const getRuleIndex = (header) => ruleHeaders.findIndex((value) => value === header);
const idxRuleId = getRuleIndex("Rule ID");
const idxStatus = getRuleIndex("상태");
const idxRider = getRuleIndex("특약명");
const idxCode = getRuleIndex("보험코드");
const idxAsIsGeneral = getRuleIndex("as-is 일반/건강 단일건");
const idxToBeGeneral = getRuleIndex("to-be 일반/건강 단일건");
const idxAsIsSimple = getRuleIndex("as-is 간편 단일건");
const idxToBeSimple = getRuleIndex("to-be 간편 단일건");
const idxMemo = getRuleIndex("검토메모");

const changeRows = ruleValues
  .slice(2)
  .filter((row) => row[idxToBeGeneral] != null && row[idxToBeGeneral] !== "" && row[idxToBeGeneral] !== row[idxAsIsGeneral]);

const rdBlob = await FileBlob.load(rdPath);
const rdWb = await SpreadsheetFile.importXlsx(rdBlob);
const rdSheet = rdWb.worksheets.getItem("Sheet1");
const rdValues = rdSheet.getUsedRange().values;
const rdHeaderNames = rdValues[1];

const rdIdxRider = 3;
const rdIdxCode = 5;
const rdIdxLimit = rdHeaderNames.findIndex((value) => value === "가입한도금액RD");
const rdIdxUnit = rdHeaderNames.findIndex((value) => value === "가입금액단위RD");
const rdIdxGuide = rdHeaderNames.findIndex((value) => value === "계약금액확인내용RD");

const historyRows = [
  ["RD판 변경 히스토리", null, null, null, null, null, null, null, null, null],
  ["변경일", "Rule ID", "상태", "보험코드", "보험명", "변경항목", "변경 전", "변경 후/가이드", "처리결과", "사유/메모"],
];

for (const row of changeRows) {
  const ruleId = row[idxRuleId];
  const status = row[idxStatus];
  const code = row[idxCode];
  const riderName = row[idxRider];
  const asIsGeneral = row[idxAsIsGeneral];
  const toBeAmount = row[idxToBeGeneral];
  const asIsSimple = row[idxAsIsSimple];
  const toBeSimple = row[idxToBeSimple] ?? row[idxToBeGeneral];
  const reviewMemo = row[idxMemo] ?? "";

  const rdRowIndex = rdValues.findIndex((rdRow, index) => index >= 4 && rdRow[rdIdxCode] === code);
  if (rdRowIndex === -1) {
    historyRows.push([
      "2026-06-16",
      ruleId,
      status,
      code,
      riderName,
      "가입한도금액RD",
      "",
      toBeAmount,
      "미반영",
      "RD판에서 동일 보험코드 행을 찾지 못함",
    ]);
    continue;
  }

  const excelRow = rdRowIndex + 1;
  const limitCell = rdSheet.getCell(excelRow - 1, rdIdxLimit);
  const unitCell = rdSheet.getCell(excelRow - 1, rdIdxUnit);
  const guideCell = rdSheet.getCell(excelRow - 1, rdIdxGuide);

  const existingLimit = rdValues[rdRowIndex][rdIdxLimit];
  const existingUnit = rdValues[rdRowIndex][rdIdxUnit];
  const isRuleValue = typeof existingLimit === "string" && existingLimit.startsWith("#");
  const isDifferentByType =
    toBeAmount != null &&
    toBeAmount !== "" &&
    toBeSimple != null &&
    toBeSimple !== "" &&
    String(toBeAmount) !== String(toBeSimple);

  if (isDifferentByType) {
    const guideText = `일반/건강(${toBeAmount})와 간편(${toBeSimple}) 한도 상이. 가입한도금액RD 룰형 반영 필요`;
    limitCell.values = [[`${existingLimit} > 룰검토필요(일반/건강=${toBeAmount}, 간편=${toBeSimple})`]];
    limitCell.format = {
      fill: fills.rule,
      font: { bold: true },
    };
    historyRows.push([
      "2026-06-16",
      ruleId,
      status,
      code,
      riderName,
      "가입한도금액RD",
      existingLimit,
      guideText,
      "룰검토필요",
      `일반/건강 as-is=${asIsGeneral}, to-be=${toBeAmount} / 간편 as-is=${asIsSimple}, to-be=${toBeSimple}. ${reviewMemo}`.trim(),
    ]);
    historyRows.push([
      "2026-06-16",
      ruleId,
      status,
      code,
      riderName,
      "가입금액단위RD",
      existingUnit,
      existingUnit,
      "유지",
      "기획 Agent 결과에 가입금액단위 변경 정보 없음",
    ]);
    continue;
  }

  if (isRuleValue) {
    const guideText = `룰형 가입한도 검토 필요: 기획 Agent 변경값 ${toBeAmount} 기준으로 ${existingLimit} 룰 조건 수정 여부 확인`;
    limitCell.values = [[`${existingLimit} > 룰검토필요(${toBeAmount})`]];
    limitCell.format = {
      fill: fills.rule,
      font: { bold: true },
    };
    guideCell.values = [[guideText]];
    guideCell.format = {
      fill: fills.rule,
      font: { bold: true },
    };
    unitCell.values = [[existingUnit]];

    historyRows.push([
      "2026-06-16",
      ruleId,
      status,
      code,
      riderName,
      "가입한도금액RD",
      existingLimit,
      guideText,
      "룰검토필요",
      `룰형 RD는 자동수정 금지. ${reviewMemo}`.trim(),
    ]);
    historyRows.push([
      "2026-06-16",
      ruleId,
      status,
      code,
      riderName,
      "가입금액단위RD",
      existingUnit,
      existingUnit,
      "유지",
      "기획 Agent 결과에 가입금액단위 변경 정보 없음",
    ]);
  } else {
    const newLimit = `${toBeAmount}~${toBeAmount}`;
    limitCell.values = [[`${existingLimit} > ${newLimit}`]];
    limitCell.format = {
      fill: fills.normal,
      font: { bold: true },
    };
    unitCell.values = [[existingUnit]];
    historyRows.push([
      "2026-06-16",
      ruleId,
      status,
      code,
      riderName,
      "가입한도금액RD",
      existingLimit,
      newLimit,
      "자동반영",
      reviewMemo,
    ]);
    historyRows.push([
      "2026-06-16",
      ruleId,
      status,
      code,
      riderName,
      "가입금액단위RD",
      existingUnit,
      existingUnit,
      "유지",
      "기획 Agent 결과에 가입금액단위 변경 정보 없음",
    ]);
  }
}

const historyName = "변경히스토리";
const existingHistory = rdWb.worksheets.items.find((sheet) => sheet.name === historyName);
if (existingHistory) {
  existingHistory.delete();
}

const historySheet = rdWb.worksheets.add(historyName);
const historyRange = `A1:J${historyRows.length}`;
historySheet.getRange(historyRange).values = historyRows;
historySheet.getRange("A1:J1").merge();
historySheet.getRange("A1:J1").format = {
  fill: fills.title,
  font: { bold: true, color: "#FFFFFF", fontSize: 14 },
  horizontalAlignment: "center",
};
historySheet.getRange("A2:J2").format = {
  fill: fills.info,
  font: { bold: true },
  horizontalAlignment: "center",
};
historySheet.getRange(historyRange).format.borders = { preset: "all", style: "thin", color: fills.border };
historySheet.getRange("A:J").format.wrapText = true;
historySheet.getRange("A:A").format.columnWidthPx = 90;
historySheet.getRange("B:B").format.columnWidthPx = 70;
historySheet.getRange("C:C").format.columnWidthPx = 80;
historySheet.getRange("D:D").format.columnWidthPx = 90;
historySheet.getRange("E:E").format.columnWidthPx = 110;
historySheet.getRange("F:F").format.columnWidthPx = 130;
historySheet.getRange("G:G").format.columnWidthPx = 120;
historySheet.getRange("H:H").format.columnWidthPx = 360;
historySheet.getRange("I:I").format.columnWidthPx = 90;
historySheet.getRange("J:J").format.columnWidthPx = 320;

historySheet.getRange("I3:I100").conditionalFormats.add("containsText", {
  text: "자동반영",
  format: { fill: "#E7F6EC", font: { bold: true, color: "#087A61" } },
});
historySheet.getRange("I3:I100").conditionalFormats.add("containsText", {
  text: "룰검토필요",
  format: { fill: fills.rule, font: { bold: true, color: "#9A3412" } },
});
historySheet.getRange("I3:I100").conditionalFormats.add("containsText", {
  text: "유지",
  format: { fill: "#F5F5F5", font: { bold: true, color: "#475467" } },
});

const inspect = await rdWb.inspect({
  kind: "table",
  sheetId: "Sheet1",
  range: "A1:O14",
  include: "values",
  tableMaxRows: 14,
  tableMaxCols: 15,
  maxChars: 10000,
});
console.log(inspect.ndjson);

const exported = await SpreadsheetFile.exportXlsx(rdWb);
await exported.save(outputPath);
console.log(outputPath);
