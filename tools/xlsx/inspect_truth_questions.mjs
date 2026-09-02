import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
if (!inputPath) {
  throw new Error("Usage: node inspect_truth_questions.mjs <workbook.xlsx>");
}

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));

const sheets = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 20000,
});
console.log("SHEETS");
console.log(sheets.ndjson);

const workbookSummary = await workbook.inspect({
  kind: "workbook,sheet,table",
  include: "id,name,range,values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 20,
  maxChars: 40000,
});
console.log("SUMMARY");
console.log(workbookSummary.ndjson);
