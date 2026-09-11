import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";


const [payloadPath, template42, output42, template43, output43] = process.argv.slice(2);
if (![payloadPath, template42, output42, template43, output43].every(Boolean)) {
  throw new Error("用法：node four_workbooks.mjs payload template42 output42 template43 output43");
}

const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));

function convertDates(rows) {
  return rows.map((row) => {
    const result = [...row];
    if (typeof result[0] === "string" && /^\d{4}-\d{2}-\d{2}$/.test(result[0])) {
      result[0] = new Date(`${result[0]}T00:00:00`);
    }
    return result;
  });
}

function repeatTemplateRows(sheet, sourceStartRow, sourceRowCount, targetRowCount, columnCount) {
  let offset = 0;
  while (offset < targetRowCount) {
    const count = Math.min(sourceRowCount, targetRowCount - offset);
    const source = sheet.getRangeByIndexes(sourceStartRow, 0, count, columnCount);
    const target = sheet.getRangeByIndexes(sourceStartRow + offset, 0, count, columnCount);
    if (offset > 0) target.copyFrom(source, "all");
    offset += count;
  }
}

function formatEmergencyBody(sheet, rowCount) {
  const body = sheet.getRangeByIndexes(1, 0, rowCount, 3);
  body.format = {
    font: { name: "宋体", size: 10, bold: false, color: "#000000" },
    borders: { preset: "all", style: "thin", color: "#000000" },
    verticalAlignment: "center",
  };
  sheet.getRangeByIndexes(1, 0, rowCount, 1).format.horizontalAlignment = "center";
  sheet.getRangeByIndexes(1, 1, rowCount, 1).format.horizontalAlignment = "right";
  sheet.getRangeByIndexes(1, 2, rowCount, 1).format.horizontalAlignment = "center";
  sheet.getRangeByIndexes(1, 0, rowCount, 1).format.numberFormat = "m/d/yy";
}

async function build42() {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(template42));
  const plan = workbook.worksheets.getItem("计划购电量");
  const storage = workbook.worksheets.getItem("充放电量");
  const emergency = workbook.worksheets.getItem("紧急购电量");

  plan.getRangeByIndexes(1, 1, payload.q42.plan.length, 146).values = payload.q42.plan;

  const storageRows = convertDates(payload.q42.storage);
  repeatTemplateRows(storage, 1, 18, storageRows.length, 6);
  storage.getRangeByIndexes(1, 0, storageRows.length, 6).values = storageRows;

  const emergencyRows = convertDates(payload.q42.emergency);
  const emergencyBody = emergency.getRangeByIndexes(1, 0, emergencyRows.length, 3);
  emergencyBody.values = emergencyRows;
  formatEmergencyBody(emergency, emergencyRows.length);

  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(output42);
}

async function build43() {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(template43));
  const plan = workbook.worksheets.getItem("计划购电量");
  const adjust = workbook.worksheets.getItem("调整购电量");
  const storage = workbook.worksheets.getItem("充放电量");
  const emergency = workbook.worksheets.getItem("紧急购电量");

  plan.getRangeByIndexes(1, 1, payload.q43.plan.length, 146).values = payload.q43.plan;
  adjust.getRangeByIndexes(1, 1, payload.q43.adjust.length, 146).values = payload.q43.adjust;

  const storageRows = convertDates(payload.q43.storage);
  repeatTemplateRows(storage, 1, 24, storageRows.length, 6);
  storage.getRangeByIndexes(1, 0, storageRows.length, 6).values = storageRows;

  const emergencyRows = convertDates(payload.q43.emergency);
  const emergencyBody = emergency.getRangeByIndexes(1, 0, emergencyRows.length, 3);
  emergencyBody.values = emergencyRows;
  formatEmergencyBody(emergency, emergencyRows.length);

  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(output43);
}

await build42();
await build43();
