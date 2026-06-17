import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const guidelinePath = String.raw`C:\Users\user\Desktop\기획agent_가이드라인v2.xlsx`;
const rdPath = String.raw`C:\Users\user\Desktop\RD판.xlsx`;
const outputDir = String.raw`C:\Users\user\Desktop\uw guide\outputs`;
const outputPath = path.join(outputDir, "RD판_지원Agent_v2_상품코드기준_차등한도반영.xlsx");

await fs.mkdir(outputDir, { recursive: true });

const fills = {
  normal: "#FFF59D",
  rule: "#F9D5A7",
  info: "#D9EAF7",
  title: "#1F4E78",
  border: "#B8C4D4",
};

const guidelineBlob = await FileBlob.load(guidelinePath);
const guidelineWb = await SpreadsheetFile.importXlsx(guidelineBlob);
const ruleSheet = guidelineWb.worksheets.getItem("Rule Master");
const ruleValues = ruleSheet.getUsedRange().values;
const headers = ruleValues[0];

const idx = (name) => headers.findIndex((value) => value === name);
const idxProductCode = idx("상품코드");
const idxRuleId = idx("Rule ID");
const idxStatus = idx("상태");
const idxRider = idx("특약명");
const idxCode = idx("보험코드");
const idxAsIsGeneral = idx("as-is 일반/건강 단일건");
const idxToBeGeneral = idx("to-be 일반/건강 단일건");
const idxAsIsSimple = idx("as-is 간편 단일건");
const idxToBeSimple = idx("to-be 간편 단일건");
const idxMemo = idx("검토메모");

const changedRows = ruleValues.slice(1).filter((row) => {
  const gChanged = row[idxToBeGeneral] != null && String(row[idxToBeGeneral]) !== String(row[idxAsIsGeneral]);
  const sChanged = row[idxToBeSimple] != null && String(row[idxToBeSimple]) !== String(row[idxAsIsSimple]);
  return gChanged || sChanged;
});

const rdBlob = await FileBlob.load(rdPath);
const rdWb = await SpreadsheetFile.importXlsx(rdBlob);
const rdSheet = rdWb.worksheets.getItem("Sheet1");
const rdValues = rdSheet.getUsedRange().values;
const rdHeaders = rdValues[1];

const rdIdxProduct = 1; // RD판의 '상품' 컬럼에 상품코드가 들어감
const rdIdxCode = 5;
const rdIdxRider = 3;
const rdIdxLimit = rdHeaders.findIndex((value) => value === "가입한도금액RD");
const rdIdxUnit = rdHeaders.findIndex((value) => value === "가입금액단위RD");

const historyRows = [
  ["RD판 변경 히스토리", null, null, null, null, null, null, null, null, null, null],
  ["변경일", "상품코드", "Rule ID", "상태", "보험코드", "보험명", "변경항목", "변경 전", "변경 후/가이드", "처리결과", "사유/메모"],
];

for (const row of changedRows) {
  const productCode = row[idxProductCode];
  const ruleId = row[idxRuleId];
  const status = row[idxStatus];
  const code = row[idxCode];
  const riderName = row[idxRider];
  const asIsGeneral = row[idxAsIsGeneral];
  const toBeGeneral = row[idxToBeGeneral] ?? row[idxAsIsGeneral];
  const asIsSimple = row[idxAsIsSimple];
  const toBeSimple = row[idxToBeSimple] ?? row[idxAsIsSimple];
  const reviewMemo = row[idxMemo] ?? "";

  const rdRowIndex = rdValues.findIndex(
    (rdRow, rowIndex) =>
      rowIndex >= 4 &&
      String(rdRow[rdIdxProduct]) === String(productCode) &&
      String(rdRow[rdIdxCode]) === String(code),
  );

  if (rdRowIndex === -1) {
    historyRows.push([
      "2026-06-16",
      productCode,
      ruleId,
      status,
      code,
      riderName,
      "가입한도금액RD",
      "",
      `일반/건강=${toBeGeneral}, 간편=${toBeSimple}`,
      "미반영",
      "상품코드+보험코드 기준 RD 행을 찾지 못함",
    ]);
    continue;
  }

  const excelRow = rdRowIndex + 1;
  const existingLimit = rdValues[rdRowIndex][rdIdxLimit];
  const existingUnit = rdValues[rdRowIndex][rdIdxUnit];
  const limitCell = rdSheet.getCell(excelRow - 1, rdIdxLimit);
  const unitCell = rdSheet.getCell(excelRow - 1, rdIdxUnit);
  const isExistingRule = typeof existingLimit === "string" && existingLimit.startsWith("#");
  const differentByType = String(toBeGeneral) !== String(toBeSimple);

  if (differentByType || isExistingRule) {
    const guideText = differentByType
      ? `일반/건강(${toBeGeneral})와 간편(${toBeSimple}) 한도 상이. 가입한도금액RD 룰형 반영 필요`
      : `기존 룰값 ${existingLimit} 검토 필요. 변경 기준 일반/건강=${toBeGeneral}, 간편=${toBeSimple}`;
    limitCell.values = [[`${existingLimit} > 룰검토필요(일반/건강=${toBeGeneral}, 간편=${toBeSimple})`]];
    limitCell.format = {
      fill: fills.rule,
      font: { bold: true },
    };
    historyRows.push([
      "2026-06-16",
      productCode,
      ruleId,
      status,
      code,
      riderName,
      "가입한도금액RD",
      existingLimit,
      guideText,
      "룰검토필요",
      `일반/건강 as-is=${asIsGeneral}, to-be=${toBeGeneral} / 간편 as-is=${asIsSimple}, to-be=${toBeSimple}. ${reviewMemo}`.trim(),
    ]);
  } else {
    const newLimit = `${toBeGeneral}~${toBeGeneral}`;
    limitCell.values = [[`${existingLimit} > ${newLimit}`]];
    limitCell.format = {
      fill: fills.normal,
      font: { bold: true },
    };
    historyRows.push([
      "2026-06-16",
      productCode,
      ruleId,
      status,
      code,
      riderName,
      "가입한도금액RD",
      existingLimit,
      newLimit,
      "자동반영",
      `일반/건강과 간편 한도 동일(${toBeGeneral}). ${reviewMemo}`.trim(),
    ]);
  }

  unitCell.values = [[existingUnit]];
  historyRows.push([
    "2026-06-16",
    productCode,
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

const historyName = "변경히스토리";
const existingHistory = rdWb.worksheets.items.find((sheet) => sheet.name === historyName);
if (existingHistory) existingHistory.delete();

const historySheet = rdWb.worksheets.add(historyName);
const historyRange = `A1:K${historyRows.length}`;
historySheet.getRange(historyRange).values = historyRows;
historySheet.getRange("A1:K1").merge();
historySheet.getRange("A1:K1").format = {
  fill: fills.title,
  font: { bold: true, color: "#FFFFFF", fontSize: 14 },
  horizontalAlignment: "center",
};
historySheet.getRange("A2:K2").format = {
  fill: fills.info,
  font: { bold: true },
  horizontalAlignment: "center",
};
historySheet.getRange(historyRange).format.borders = { preset: "all", style: "thin", color: fills.border };
historySheet.getRange("A:K").format.wrapText = true;
historySheet.getRange("A:A").format.columnWidthPx = 90;
historySheet.getRange("B:B").format.columnWidthPx = 90;
historySheet.getRange("C:C").format.columnWidthPx = 70;
historySheet.getRange("D:D").format.columnWidthPx = 70;
historySheet.getRange("E:E").format.columnWidthPx = 90;
historySheet.getRange("F:F").format.columnWidthPx = 110;
historySheet.getRange("G:G").format.columnWidthPx = 130;
historySheet.getRange("H:H").format.columnWidthPx = 120;
historySheet.getRange("I:I").format.columnWidthPx = 330;
historySheet.getRange("J:J").format.columnWidthPx = 90;
historySheet.getRange("K:K").format.columnWidthPx = 360;

historySheet.getRange("J3:J100").conditionalFormats.add("containsText", {
  text: "자동반영",
  format: { fill: "#E7F6EC", font: { bold: true, color: "#087A61" } },
});
historySheet.getRange("J3:J100").conditionalFormats.add("containsText", {
  text: "룰검토필요",
  format: { fill: fills.rule, font: { bold: true, color: "#9A3412" } },
});

const exported = await SpreadsheetFile.exportXlsx(rdWb);
await exported.save(outputPath);
console.log(outputPath);
