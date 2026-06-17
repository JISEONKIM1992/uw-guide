import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const guidelinePath = String.raw`C:\Users\user\Desktop\신계약가이드라인_공통코어_샘플결과_v2.xlsx`;
const rdPath = String.raw`C:\Users\user\Desktop\RD판.xlsx`;
const outputDir = String.raw`C:\Users\user\Desktop\uw guide\outputs`;
const outputPath = path.join(outputDir, "RD판_지원Agent_샘플수정.xlsx");

await fs.mkdir(outputDir, { recursive: true });

const guidelineBlob = await FileBlob.load(guidelinePath);
const guidelineWb = await SpreadsheetFile.importXlsx(guidelineBlob);
const ruleMaster = guidelineWb.worksheets.getItem("Rule Master");

const rdBlob = await FileBlob.load(rdPath);
const rdWb = await SpreadsheetFile.importXlsx(rdBlob);
const rdSheet = rdWb.worksheets.getItem("Sheet1");

const ruleRange = ruleMaster.getRange("A1:R12");
const ruleValues = ruleRange.values;
const ruleHeaders = ruleValues[1];

const headerIndex = (name) => ruleHeaders.findIndex((value) => value === name);
const statusCol = headerIndex("상태");
const riderCol = headerIndex("특약명");
const codeCol = headerIndex("보험코드");
const toBeGeneralCol = headerIndex("to-be 일반/건강 단일건");

const targetRule = ruleValues
  .slice(2)
  .find(
    (row) =>
      row[statusCol] === "시행예정" &&
      row[riderCol] === "소액암진단" &&
      row[codeCol] === "LI00113",
  );

if (!targetRule) {
  throw new Error("기획 Agent 결과에서 소액암진단 시행예정 건을 찾지 못했습니다.");
}

const targetAmount = targetRule[toBeGeneralCol];
if (!targetAmount) {
  throw new Error("기획 Agent 결과에 변경 후 단일건 값이 없습니다.");
}

const rdValues = rdSheet.getRange("A1:O14").values;
const rdHeaders = rdValues[1];
const limitCol = rdHeaders.findIndex((value) => value === "가입한도금액RD");
const unitCol = rdHeaders.findIndex((value) => value === "가입금액단위RD");
const insuranceCodeCol = 5;

const targetRowIndex = rdValues.findIndex(
  (row, index) => index >= 4 && row[insuranceCodeCol] === "LI00113",
);

if (targetRowIndex === -1) {
  throw new Error("RD판에서 보험코드 LI00113 행을 찾지 못했습니다.");
}

const excelRow = targetRowIndex + 1;
const limitCell = rdSheet.getCell(excelRow - 1, limitCol);
const beforeLimit = rdValues[targetRowIndex][limitCol];
const afterLimit = `${targetAmount}~${targetAmount}`;
limitCell.values = [[`${beforeLimit} > ${afterLimit}`]];
limitCell.format = {
  fill: "#FFF59D",
  font: { bold: true },
};

// 이번 샘플은 가이드라인 결과에 단위 변경 컬럼이 없어 기존 값을 유지한다.
const currentUnit =
  rdSheet.getCell(excelRow - 1, unitCol).values?.[0]?.[0] ?? rdValues[targetRowIndex][unitCol];
rdSheet.getCell(excelRow - 1, unitCol).values = [[currentUnit]];

const historySheetName = "변경히스토리";
const existingHistory = rdWb.worksheets.items.find((sheet) => sheet.name === historySheetName);
if (existingHistory) {
  existingHistory.delete();
}

const historySheet = rdWb.worksheets.add(historySheetName);
historySheet.getRange("A1:H4").values = [
  ["RD판 변경 히스토리", null, null, null, null, null, null, null],
  ["변경일", "구분", "보험코드", "보험명", "변경항목", "변경 전", "변경 후", "변경사유"],
  [
    "2026-06-15",
    "특약",
    "LI00113",
    "소액암진단",
    "가입한도금액RD",
    beforeLimit,
    afterLimit,
    "기획 Agent 결과 기준 소액암진단 단일건 1000 -> 2000 변경 반영",
  ],
  [
    "2026-06-15",
    "특약",
    "LI00113",
    "소액암진단",
    "가입금액단위RD",
    currentUnit,
    currentUnit,
    "기획 Agent 결과에 단위 변경 정보 없음. 기존값 유지",
  ],
];

historySheet.getRange("A1:H1").merge();
historySheet.getRange("A1:H1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", fontSize: 14 },
  horizontalAlignment: "center",
};
historySheet.getRange("A2:H2").format = {
  fill: "#D9EAF7",
  font: { bold: true },
  horizontalAlignment: "center",
};
historySheet.getRange("A1:H4").format.borders = { preset: "all", style: "thin", color: "#B8C4D4" };
historySheet.getRange("F3:G3").format = {
  fill: "#FFF59D",
  font: { bold: true },
};
historySheet.getRange("F4:G4").format = {
  fill: "#F5F5F5",
};
historySheet.getRange("A:H").format.wrapText = true;
historySheet.getRange("A:A").format.columnWidthPx = 90;
historySheet.getRange("B:B").format.columnWidthPx = 70;
historySheet.getRange("C:C").format.columnWidthPx = 90;
historySheet.getRange("D:D").format.columnWidthPx = 110;
historySheet.getRange("E:E").format.columnWidthPx = 130;
historySheet.getRange("F:G").format.columnWidthPx = 120;
historySheet.getRange("H:H").format.columnWidthPx = 320;

const inspect = await rdWb.inspect({
  kind: "table",
  sheetId: "Sheet1",
  range: `A${excelRow}:O${excelRow}`,
  include: "values",
  tableMaxRows: 1,
  tableMaxCols: 15,
  maxChars: 3000,
});
console.log(inspect.ndjson);

const exported = await SpreadsheetFile.exportXlsx(rdWb);
await exported.save(outputPath);
console.log(outputPath);
