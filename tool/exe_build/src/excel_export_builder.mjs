import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = process.argv[2];
if (!inputPath) {
  throw new Error("Missing input JSON path");
}

const raw = await fs.readFile(inputPath, "utf8");
const payload = JSON.parse(raw.replace(/^\uFEFF/, ""));
const rows = Array.isArray(payload.rows) ? payload.rows : [];
const outputPath = payload.output_path;
if (!outputPath) {
  throw new Error("Missing output_path");
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("文案汇总");

sheet.getRange("A1:F1").values = [[
  "sku",
  "货号",
  "标题",
  "简介",
  "json富文本内容",
  "是否上传成功",
]];

const values = rows.map((row) => [
  row.sku ?? "",
  row.product_code ?? "",
  row.title ?? "",
  row.summary ?? "",
  row.rich_json ?? "",
  row.upload_status ?? "",
]);

if (values.length > 0) {
  sheet.getRangeByIndexes(1, 0, values.length, 6).values = values;
}

sheet.getRange("A1:F1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
};

const widths = [120, 140, 260, 300, 520, 140];
for (let i = 0; i < widths.length; i += 1) {
  const range = sheet.getRangeByIndexes(0, i, Math.max(values.length + 1, 2), 1);
  range.format.columnWidthPx = widths[i];
  range.format.wrapText = true;
}

sheet.freezePanes.freezeRows(1);

const lastRow = Math.max(values.length + 1, 2);
sheet.getRange(`A1:F${lastRow}`).format.borders = {
  top: { style: "Continuous", color: "#D9E2F3" },
  bottom: { style: "Continuous", color: "#D9E2F3" },
  left: { style: "Continuous", color: "#D9E2F3" },
  right: { style: "Continuous", color: "#D9E2F3" },
  insideHorizontal: { style: "Continuous", color: "#EDEDED" },
  insideVertical: { style: "Continuous", color: "#EDEDED" },
};

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
