import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = "D:/ozon/tool/ozon_batch_upload_template.xlsx";

const workbook = Workbook.create();
const input = workbook.worksheets.add("上架填写");
const guide = workbook.worksheets.add("填写说明");

input.getRange("A1:C1").values = [["货号", "标题", "简介"]];
input.getRange("A2:C5").values = [
  ["SKU001", "Женский платок квадратный с геометрическим принтом", "Легкий женский платок для повседневного образа. Подходит для прогулок, поездок и сочетания с разной одеждой."],
  ["SKU002", "Женский шарф мягкий однотонный для повседневного ношения", "Мягкий аксессуар для базового гардероба. Можно использовать как шарф, накидку или декоративный элемент образа."],
  ["", "", ""],
  ["", "", ""],
];
input.getRange("A1:C1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
};
input.getRange("A1:C200").format = {
  borders: {
    top: { style: "Continuous", color: "#D9E2F3" },
    bottom: { style: "Continuous", color: "#D9E2F3" },
    left: { style: "Continuous", color: "#D9E2F3" },
    right: { style: "Continuous", color: "#D9E2F3" },
    insideHorizontal: { style: "Continuous", color: "#EDEDED" },
    insideVertical: { style: "Continuous", color: "#EDEDED" },
  },
  wrapText: true,
};
input.getRange("A:A").format.columnWidthPx = 160;
input.getRange("B:B").format.columnWidthPx = 420;
input.getRange("C:C").format.columnWidthPx = 680;
input.getRange("A2:C200").format.rowHeightPx = 58;
input.freezePanes.freezeRows(1);

guide.getRange("A1:B1").values = [["字段", "说明"]];
guide.getRange("A2:B7").values = [
  ["货号", "必须和 3:4 输出目录下的文件夹名完全一致，例如文件夹 SKU001 对应货号 SKU001。"],
  ["标题", "上传到 Ozon 的商品标题。建议使用俄文，不能为空。"],
  ["简介", "上传到 Ozon 的商品描述/简介。建议使用俄文，不能为空。"],
  ["图片", "不用填在 Excel 里。程序会自动读取 3:4 输出目录中同名货号文件夹里的图片。"],
  ["表头", "第一行必须保持为：货号、标题、简介。不要改名，不要合并单元格。"],
  ["保存", "填写后保存为 .xlsx，再在工具的“上架 Excel”中选择该文件。"],
];
guide.getRange("A1:B1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
};
guide.getRange("A1:B7").format = {
  borders: {
    top: { style: "Continuous", color: "#D9E2F3" },
    bottom: { style: "Continuous", color: "#D9E2F3" },
    left: { style: "Continuous", color: "#D9E2F3" },
    right: { style: "Continuous", color: "#D9E2F3" },
    insideHorizontal: { style: "Continuous", color: "#EDEDED" },
    insideVertical: { style: "Continuous", color: "#EDEDED" },
  },
  wrapText: true,
};
guide.getRange("A:A").format.columnWidthPx = 140;
guide.getRange("B:B").format.columnWidthPx = 760;
guide.getRange("A2:B7").format.rowHeightPx = 52;
guide.freezePanes.freezeRows(1);

await fs.mkdir("D:/ozon/tool", { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
