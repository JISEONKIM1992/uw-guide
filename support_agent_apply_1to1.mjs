import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const guidelinePath = String.raw`C:\Users\user\Desktop\신계약가이드라인_공통코어_샘플결과_v2.xlsx`;
const rdPath = String.raw`C:\Users\user\Desktop\RD판.xlsx`;
const outputDir = String.raw`C:\Users\user\Desktop\uw guide\outputs`;
const outputPath = path.join(outputDir, "RD판_지원Agent_1대1매핑_자동수정.xlsx");

await fs.mkdir(outputDir, { recursive: true });

const yellowFill = "#FFF59D";
const blueFill = "#D9EAF7";
const titleFill = "#1F4E78";
const borderColor = "#B8C4D4";

const guidelineBlob = await FileBlob.load(guidelinePath);
const guidelineWb = await SpreadsheetFile.importXlsx(guidelineBlob);
const ruleSheet = guidelineWb.worksheets.getItem("Rule Master");
const ruleValues = ruleSheet.getUsedRange().values;
const ruleHeaders = ruleValues[1];

const getRuleIndex = (header) => ruleHeaders.findIndex((value) => value === header);
const idxStatus = getRuleIndex("상태");
const idxRuleId = getRuleIndex("Rule ID");
const idxRider = getRuleIndex("특약명");
const idxCode = getRuleIndex("보험코드");
const idxAsIsGeneral = getRuleIndex("as-is 일반/건강 단일건");
const idxToBeGeneral = getRuleIndex("to-be 일반/건강 단일건");
const idxMemo = getRuleIndex("검토메모");

const changeRows = ruleValues
  .slice(2)
  .filter((row) => row[idxStatus] === "시행예정" && row[idxToBeGeneral] != null && row[idxToBeGeneral] !== "");

const rdBlob = await FileBlob.load(rdPath);
const rdWb = await SpreadsheetFile.importXlsx(rdBlob);
const rdSheet = rdWb.worksheets.getItem("Sheet1");
const rdValues = rdSheet.getUsedRange().values;
const rdHeaderNames = rdValues[1];

const rdIdxCode = 5;
const rdIdxRider = 3;
const rdIdxLimit = rdHeaderNames.findIndex((value) => value === "가입한도금액RD");
const rdIdxUnit = rdHeaderNames.findIndex((value) => value === "가입금액단위RD");

const historyRows = [
  ["RD판 변경 히스토리", null, null, null, null, null, null, null, null],
  ["변경일", "Rule ID", "보험코드", "보험명", "변경항목", "변경 전", "변경 후", "처리결과", "사유/메모"],
];

for (const row of changeRows) {
  const code = row[idxCode];
  const riderName = row[idxRider];
  const beforeAmount = row[idxAsIsGeneral];
  const afterAmount = row[idxToBeGeneral];
  const reviewMemo = row[idxMemo] ?? "";

  const rdRowIndex = rdValues.findIndex(
    (rdRow, index) => index >= 4 && rdRow[rdIdxCode] === code,
  );

  if (rdRowIndex === -1) {
    historyRows.push([
      "2026-06-16",
      row[idxRuleId],
      code,
      riderName,
      "가입한도금액RD",
      beforeAmount,
      afterAmount,
      "미반영",
      "RD판에서 동일 보험코드 행을 찾지 못함",
    ]);
    continue;
  }

  const excelRow = rdRowIndex + 1;
  const limitCell = rdSheet.getCell(excelRow - 1, rdIdxLimit);
  const existingLimit = rdValues[rdRowIndex][rdIdxLimit];
  const newLimit = `${afterAmount}~${afterAmount}`;
  limitCell.values = [[`${existingLimit} > ${newLimit}`]];
  limitCell.format = {
    fill: yellowFill,
    font: { bold: true },
  };

  const unitCell = rdSheet.getCell(excelRow - 1, rdIdxUnit);
  const existingUnit = rdValues[rdRowIndex][rdIdxUnit];
  unitCell.values = [[existingUnit]];

  historyRows.push([
    "2026-06-16",
    row[idxRuleId],
    code,
    riderName,
    "가입한도금액RD",
    existingLimit,
    newLimit,
    "자동반영",
    `특약코드=템플릿ID 1:1 매핑 전제. ${reviewMemo}`.trim(),
  ]);

  historyRows.push([
    "2026-06-16",
    row[idxRuleId],
    code,
    riderName,
    "가입금액단위RD",
    existingUnit,
    existingUnit,
    "유지",
    "기획 Agent 결과에 가입금액단위 변경 정보 없음",
  ]);
}

const historyName = "변경히스토리";
const existingHistory = rdWb.worksheets.items.find((sheet) => sheet.name === historyName);
if (existingHistory) {
  existingHistory.delete();
}

const historySheet = rdWb.worksheets.add(historyName);
const historyRange = `A1:I${historyRows.length}`;
historySheet.getRange(historyRange).values = historyRows;
historySheet.getRange("A1:I1").merge();
historySheet.getRange("A1:I1").format = {
  fill: titleFill,
  font: { bold: true, color: "#FFFFFF", fontSize: 14 },
  horizontalAlignment: "center",
};
historySheet.getRange("A2:I2").format = {
  fill: blueFill,
  font: { bold: true },
  horizontalAlignment: "center",
};
historySheet.getRange(historyRange).format.borders = { preset: "all", style: "thin", color: borderColor };
if (historyRows.length > 2) {
  historySheet.getRange(`F3:G${historyRows.length}`).format = {
    fill: "#FFFDF2",
  };
}
historySheet.getRange("H3:H100").conditionalFormats.add("containsText", {
  text: "자동반영",
  format: { fill: "#E7F6EC", font: { bold: true, color: "#087A61" } },
});
historySheet.getRange("H3:H100").conditionalFormats.add("containsText", {
  text: "미반영",
  format: { fill: "#FDECEC", font: { bold: true, color: "#B42318" } },
});
historySheet.getRange("A:I").format.wrapText = true;
historySheet.getRange("A:A").format.columnWidthPx = 90;
historySheet.getRange("B:B").format.columnWidthPx = 70;
historySheet.getRange("C:C").format.columnWidthPx = 90;
historySheet.getRange("D:D").format.columnWidthPx = 110;
historySheet.getRange("E:E").format.columnWidthPx = 130;
historySheet.getRange("F:G").format.columnWidthPx = 120;
historySheet.getRange("H:H").format.columnWidthPx = 90;
historySheet.getRange("I:I").format.columnWidthPx = 360;

const inspect = await rdWb.inspect({
  kind: "table",
  sheetId: "Sheet1",
  range: "A1:O14",
  include: "values",
  tableMaxRows: 14,
  tableMaxCols: 15,
  maxChars: 8000,
});
console.log(inspect.ndjson);

const exported = await SpreadsheetFile.exportXlsx(rdWb);
await exported.save(outputPath);
console.log(outputPath);
