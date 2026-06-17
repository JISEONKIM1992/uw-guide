import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = String.raw`C:\Users\user\Desktop\uw guide\outputs\RD판_지원Agent_v2_상품코드기준_차등한도반영.xlsx`;
const outputDir = String.raw`C:\Users\user\Desktop\uw guide\outputs`;
const outputPath = path.join(outputDir, "RD판_지원Agent_v2_요약시트추가.xlsx");

await fs.mkdir(outputDir, { recursive: true });

const blob = await FileBlob.load(sourcePath);
const wb = await SpreadsheetFile.importXlsx(blob);

const oldTable = wb.worksheets.items.find((sheet) => sheet.name === "지원Agent_수정필요항목");
if (oldTable) oldTable.delete();
const oldPage = wb.worksheets.items.find((sheet) => sheet.name === "지원Agent_1장요약");
if (oldPage) oldPage.delete();

const historySheet = wb.worksheets.getItem("변경히스토리");
const historyValues = historySheet.getUsedRange().values;
const dataRows = historyValues.slice(2).filter((row) => row[6] === "가입한도금액RD");

const tableSheet = wb.worksheets.add("지원Agent_수정필요항목");
const tableRows = [
  ["지원 Agent RD 수정 필요 항목", null, null, null, null, null, null, null],
  ["상품코드", "Rule ID", "보험코드", "보험명", "RD 대상항목", "처리구분", "변경 후/가이드", "사유/메모"],
];

for (const row of dataRows) {
  tableRows.push([
    row[1],
    row[2],
    row[4],
    row[5],
    row[6],
    row[9],
    row[8],
    row[10],
  ]);
}

tableSheet.getRange(`A1:H${tableRows.length}`).values = tableRows;
tableSheet.getRange("A1:H1").merge();
tableSheet.getRange("A1:H1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", fontSize: 14 },
  horizontalAlignment: "center",
};
tableSheet.getRange("A2:H2").format = {
  fill: "#D9EAF7",
  font: { bold: true },
  horizontalAlignment: "center",
};
tableSheet.getRange(`A1:H${tableRows.length}`).format.borders = { preset: "all", style: "thin", color: "#B8C4D4" };
tableSheet.getRange("A:H").format.wrapText = true;
tableSheet.getRange("A:A").format.columnWidthPx = 90;
tableSheet.getRange("B:B").format.columnWidthPx = 70;
tableSheet.getRange("C:C").format.columnWidthPx = 90;
tableSheet.getRange("D:D").format.columnWidthPx = 110;
tableSheet.getRange("E:E").format.columnWidthPx = 130;
tableSheet.getRange("F:F").format.columnWidthPx = 100;
tableSheet.getRange("G:G").format.columnWidthPx = 300;
tableSheet.getRange("H:H").format.columnWidthPx = 340;
tableSheet.getRange(`F3:F${tableRows.length}`).conditionalFormats.add("containsText", {
  text: "자동반영",
  format: { fill: "#E7F6EC", font: { bold: true, color: "#087A61" } },
});
tableSheet.getRange(`F3:F${tableRows.length}`).conditionalFormats.add("containsText", {
  text: "룰검토필요",
  format: { fill: "#F9D5A7", font: { bold: true, color: "#9A3412" } },
});

const pageSheet = wb.worksheets.add("지원Agent_1장요약");
pageSheet.showGridLines = false;

pageSheet.getRange("A1:J2").merge();
pageSheet.getRange("A1:J2").values = [["지원 Agent RD 수정 필요 항목 1장 요약"]];
pageSheet.getRange("A1:J2").format = {
  fill: "#16324F",
  font: { bold: true, color: "#FFFFFF", fontSize: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};

pageSheet.getRange("A4:E6").merge();
pageSheet.getRange("A4:E6").values = [[`자동반영후보\n${dataRows.filter((r) => r[9] === "자동반영").length}건`]];
pageSheet.getRange("A4:E6").format = {
  fill: "#E7F6EC",
  font: { bold: true, color: "#087A61", fontSize: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#B8C4D4" },
};

pageSheet.getRange("F4:J6").merge();
pageSheet.getRange("F4:J6").values = [[`룰검토필요\n${dataRows.filter((r) => r[9] === "룰검토필요").length}건`]];
pageSheet.getRange("F4:J6").format = {
  fill: "#F9D5A7",
  font: { bold: true, color: "#9A3412", fontSize: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#B8C4D4" },
};

let currentRow = 8;
for (const row of dataRows) {
  const statusColor = row[9] === "자동반영" ? "#E7F6EC" : "#FFF1E6";
  const title = `[${row[9]}] ${row[5]} (${row[4]})`;
  const body = `상품코드: ${row[1]}\n변경항목: ${row[6]}\n제안: ${row[8]}\n메모: ${row[10]}`;
  pageSheet.getRange(`A${currentRow}:J${currentRow + 3}`).merge();
  pageSheet.getRange(`A${currentRow}:J${currentRow + 3}`).values = [[`${title}\n${body}`]];
  pageSheet.getRange(`A${currentRow}:J${currentRow + 3}`).format = {
    fill: statusColor,
    font: { bold: true, color: "#17202E", fontSize: 12 },
    horizontalAlignment: "left",
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#B8C4D4" },
  };
  currentRow += 5;
}

pageSheet.getRange("A:J").format.columnWidthPx = 95;
pageSheet.getRange("1:40").format.rowHeightPx = 24;

const exported = await SpreadsheetFile.exportXlsx(wb);
await exported.save(outputPath);
console.log(outputPath);
