import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const guidelinePath = String.raw`C:\Users\user\Desktop\기획agent_가이드라인v2.xlsx`;
const outputDir = String.raw`C:\Users\user\Desktop\uw guide\outputs`;
const outputPath = path.join(outputDir, "기획agent_가이드라인v2_지원Agent정리.xlsx");

await fs.mkdir(outputDir, { recursive: true });

const guidelineBlob = await FileBlob.load(guidelinePath);
const wb = await SpreadsheetFile.importXlsx(guidelineBlob);
const ruleSheet = wb.worksheets.getItem("Rule Master");
const values = ruleSheet.getUsedRange().values;
const headers = values[0];

const idx = (name) => headers.findIndex((value) => value === name);
const iProduct = idx("상품코드");
const iRuleId = idx("Rule ID");
const iStatus = idx("상태");
const iRider = idx("특약명");
const iCode = idx("보험코드");
const iAsIsGeneral = idx("as-is 일반/건강 단일건");
const iToBeGeneral = idx("to-be 일반/건강 단일건");
const iAsIsSimple = idx("as-is 간편 단일건");
const iToBeSimple = idx("to-be 간편 단일건");
const iMemo = idx("검토메모");

const rows = values.slice(1).filter((row) => {
  const gChanged = String(row[iAsIsGeneral] ?? "") !== String(row[iToBeGeneral] ?? row[iAsIsGeneral] ?? "");
  const sChanged = String(row[iAsIsSimple] ?? "") !== String(row[iToBeSimple] ?? row[iAsIsSimple] ?? "");
  return gChanged || sChanged;
});

const sheetName = "지원Agent_수정필요항목";
const existing = wb.worksheets.items.find((sheet) => sheet.name === sheetName);
if (existing) existing.delete();
const sheet = wb.worksheets.add(sheetName);

const out = [
  ["지원 Agent RD 수정 필요 항목", null, null, null, null, null, null, null, null, null, null, null],
  ["상품코드", "Rule ID", "상태", "특약명", "보험코드", "일반/건강(as-is)", "일반/건강(to-be)", "간편(as-is)", "간편(to-be)", "RD판 대상항목", "처리구분", "지원Agent 제안"],
];

for (const row of rows) {
  const gAsIs = row[iAsIsGeneral];
  const gToBe = row[iToBeGeneral] ?? gAsIs;
  const sAsIs = row[iAsIsSimple];
  const sToBe = row[iToBeSimple] ?? sAsIs;
  const differentByType = String(gToBe) !== String(sToBe);
  const treatment = differentByType ? "룰검토필요" : "자동반영후보";
  const proposal = differentByType
    ? `가입한도금액RD는 룰형 반영 필요 (일반/건강=${gToBe}, 간편=${sToBe})`
    : `가입한도금액RD = ${gToBe}~${gToBe}`;

  out.push([
    row[iProduct],
    row[iRuleId],
    row[iStatus],
    row[iRider],
    row[iCode],
    gAsIs,
    gToBe,
    sAsIs,
    sToBe,
    "가입한도금액RD",
    treatment,
    proposal,
  ]);

  out.push([
    row[iProduct],
    row[iRuleId],
    row[iStatus],
    row[iRider],
    row[iCode],
    gAsIs,
    gToBe,
    sAsIs,
    sToBe,
    "검토메모",
    "참고",
    row[iMemo] ?? "",
  ]);
}

const range = `A1:L${out.length}`;
sheet.getRange(range).values = out;
sheet.getRange("A1:L1").merge();
sheet.getRange("A1:L1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", fontSize: 14 },
  horizontalAlignment: "center",
};
sheet.getRange("A2:L2").format = {
  fill: "#D9EAF7",
  font: { bold: true },
  horizontalAlignment: "center",
};
sheet.getRange(range).format.borders = { preset: "all", style: "thin", color: "#B8C4D4" };
sheet.getRange("A:L").format.wrapText = true;
sheet.getRange("A:A").format.columnWidthPx = 90;
sheet.getRange("B:B").format.columnWidthPx = 70;
sheet.getRange("C:C").format.columnWidthPx = 70;
sheet.getRange("D:D").format.columnWidthPx = 110;
sheet.getRange("E:E").format.columnWidthPx = 90;
sheet.getRange("F:I").format.columnWidthPx = 100;
sheet.getRange("J:J").format.columnWidthPx = 130;
sheet.getRange("K:K").format.columnWidthPx = 100;
sheet.getRange("L:L").format.columnWidthPx = 300;
sheet.getRange(`K3:K${out.length}`).conditionalFormats.add("containsText", {
  text: "자동반영후보",
  format: { fill: "#E7F6EC", font: { bold: true, color: "#087A61" } },
});
sheet.getRange(`K3:K${out.length}`).conditionalFormats.add("containsText", {
  text: "룰검토필요",
  format: { fill: "#F9D5A7", font: { bold: true, color: "#9A3412" } },
});
sheet.getRange(`K3:K${out.length}`).conditionalFormats.add("containsText", {
  text: "참고",
  format: { fill: "#F5F5F5", font: { bold: true, color: "#475467" } },
});

const exported = await SpreadsheetFile.exportXlsx(wb);
await exported.save(outputPath);
console.log(outputPath);
