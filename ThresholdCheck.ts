function main(workbook: ExcelScript.Workbook) {
  const configSheetName = "Config";
  const realThresholdColumn = "M"; // decoded (decimal) "Threshold Comp" written by Measurements.ts, raw firmware value for every row
  const calcThresholdColumn = "S"; // "Compensated threshold" written by Measurements.ts, only filled inside an active S-block
  const dataStartRow = 2; // row 1 holds the input column labels, real data starts at row 2

  const outputColumn = "V";
  const outputHeader = "Threshold check";

  const configSheet = workbook.getWorksheet(configSheetName);
  if (!configSheet) {
    console.log(`Worksheet "${configSheetName}" not found.`);
    return;
  }

  const sheetConfigs = parseConfig(configSheet);

  for (const cfg of sheetConfigs) {
    if (!cfg.generate) {
      console.log(`Skipping "${cfg.sheetName}": generate != 1.`);
      continue;
    }

    if (!cfg.dataWithThreshold) {
      console.log(`Skipping "${cfg.sheetName}": data with threshold != 1 (no real Threshold Comp to compare).`);
      continue;
    }

    const sheet = workbook.getWorksheet(cfg.sheetName);
    if (!sheet) {
      console.log(`Skipping "${cfg.sheetName}": worksheet does not exist.`);
      continue;
    }

    const startRow = dataStartRow;
    const usedRange = sheet.getUsedRange();
    const lastRow = usedRange.getRowIndex() + usedRange.getRowCount();
    if (lastRow < startRow) {
      continue;
    }

    sheet.getRange(`${outputColumn}1:${outputColumn}${lastRow}`).clear();

    const realValues = sheet.getRange(`${realThresholdColumn}${startRow}:${realThresholdColumn}${lastRow}`).getValues();
    const calcValues = sheet.getRange(`${calcThresholdColumn}${startRow}:${calcThresholdColumn}${lastRow}`).getValues();

    // Only compared where Measurements.ts has a calculated threshold (i.e. inside an active S-block); blank otherwise.
    const output: (number | string)[][] = [];
    for (let i = 0; i < realValues.length; i++) {
      const calc = calcValues[i][0];
      output.push([calc === "" ? "" : (realValues[i][0] as number)]);
    }

    sheet.getRange(`${outputColumn}1`).setValue(outputHeader);
    sheet.getRange(`${outputColumn}${startRow}:${outputColumn}${lastRow}`).setValues(output);

    formatGeneratedColumn(sheet, outputColumn, startRow, lastRow);

    for (let i = 0; i < realValues.length; i++) {
      const calc = calcValues[i][0];
      if (calc === "") {
        continue;
      }

      const match = (realValues[i][0] as number) === (calc as number);
      const row = startRow + i;
      sheet.getRange(`${outputColumn}${row}`).getFormat().getFill().setColor(match ? "#C6EFCE" : "#FFC7CE");
    }

    console.log(`"${cfg.sheetName}": updated column ${outputColumn} (threshold check), rows ${startRow}-${lastRow}.`);
  }
}

interface SheetConfig {
  sheetName: string;
  generate: boolean;
  dataWithThreshold: boolean;
  sensitivity: number;
  slopeThreshold: number;
  resetThreshold: number;
  l0Sensitivity: number;
}

// Bold + light-filled header; subtle alternating row shading on the data; thin grid borders; centered; column auto-fit to the data, then widened if needed so header words don't split mid-word.
function formatGeneratedColumn(sheet: ExcelScript.Worksheet, column: string, startRow: number, lastRow: number): void {
  const headerRange = sheet.getRange(`${column}1`);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().setWrapText(true);
  headerRange.getFormat().getFill().setColor("#D9E1F2");

  const dataRange = sheet.getRange(`${column}${startRow}:${column}${lastRow}`);
  dataRange.getFormat().getFill().clear();
  applyRowBanding(sheet, column, startRow, lastRow);

  applyBorders(sheet.getRange(`${column}1:${column}${lastRow}`));
  sheet.getRange(`${column}1:${column}${lastRow}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  sheet.getRange(`${column}:${column}`).getFormat().autofitColumns();
  ensureHeaderColumnWidth(sheet, column);
  headerRange.getFormat().autofitRows();
}

// Very light shading on every other data row (header excluded) so consecutive rows are easier to tell apart.
function applyRowBanding(sheet: ExcelScript.Worksheet, column: string, startRow: number, lastRow: number): void {
  for (let row = startRow + 1; row <= lastRow; row += 2) {
    sheet.getRange(`${column}${row}`).getFormat().getFill().setColor("#F2F2F2");
  }
}

// Widens the column if its auto-fit width would force the header's longest word to split across lines.
function ensureHeaderColumnWidth(sheet: ExcelScript.Worksheet, column: string): void {
  const header = String(sheet.getRange(`${column}1`).getValue());
  const longestWord = header.split(" ").reduce((max, word) => Math.max(max, word.length), 0);
  const minWidth = longestWord * 10 + 20; // generous per-character estimate (headers are bold), just enough to avoid mid-word wraps
  const columnRange = sheet.getRange(`${column}:${column}`);
  if (columnRange.getFormat().getColumnWidth() < minWidth) {
    columnRange.getFormat().setColumnWidth(minWidth);
  }
}

// Thin continuous borders around every cell (outer edges + inner grid lines) in the given range.
function applyBorders(range: ExcelScript.Range): void {
  const edges: ExcelScript.BorderIndex[] = [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical
  ];
  for (const edge of edges) {
    const border = range.getFormat().getRangeBorder(edge);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
  }
}

// Reads the Config sheet: one row per measurement sheet. Layout: A Measurements (sheet name), B generate, C data with threshold, D Sensitivity, E Slope Threshold, F Reset Threshold, G L0 Sensitivity.
function parseConfig(configSheet: ExcelScript.Worksheet): SheetConfig[] {
  const values = configSheet.getUsedRange().getValues();

  const sheetConfigs: SheetConfig[] = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sheetNameCell = row[0];
    if (!sheetNameCell || String(sheetNameCell).trim() === "") {
      continue;
    }

    sheetConfigs.push({
      sheetName: String(sheetNameCell).trim(),
      generate: Number(row[1]) === 1,
      dataWithThreshold: Number(row[2]) === 1,
      sensitivity: Number(row[3]),
      slopeThreshold: Number(row[4]),
      resetThreshold: Number(row[5]),
      l0Sensitivity: Number(row[6])
    });
  }

  return sheetConfigs;
}
