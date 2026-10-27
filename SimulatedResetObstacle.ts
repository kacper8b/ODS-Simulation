function main(workbook: ExcelScript.Workbook) {
  const configSheetName = "Config";
  const ucapFiltColumn = "K"; // decoded (decimal) columns written by Measurements.ts: Line, Ucap, Ucap Comp, Ucap Filt, Symbol, Threshold Comp = H-M
  const detectionColumn = "L";
  const compThresholdColumn = "S"; // written by Measurements.ts ("Compensated threshold")
  const movementColumn = "O"; // written by Measurements.ts ("Mvt"), gates S detection below
  const dataStartRow = 2; // row 1 holds the input column labels, real data starts at row 2
  const ucapTableLength = 150; // firmware: UCAP_TABLE_LENGTH, cnt counts down from this value to 0
  const prevSamples = 10; // firmware: GetnPreviosPosition(10u)
  const ucapStabilisation = 25; // firmware: UCAP_STABILISATION - detection doesn't run until this many samples after the movement starts

  const simOutputColumn = "T"; // merged S/R/E simulation
  const simOutputHeader = "S/R/E detect (sim)";
  const simValueColumn = "U"; // S: saved baseline (max); R/E: "value (v1 - v2)" showing the calculation that decided the marker
  const simValueHeader = "S/R/E value (sim)";

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

    sheet.getRange(`${simOutputColumn}1:${simValueColumn}${lastRow}`).clear();

    const ucapFiltValues = sheet.getRange(`${ucapFiltColumn}${startRow}:${ucapFiltColumn}${lastRow}`).getValues();
    const infoValues = sheet.getRange(`${detectionColumn}${startRow}:${detectionColumn}${lastRow}`).getValues();
    const compThresholdValues = sheet.getRange(`${compThresholdColumn}${startRow}:${compThresholdColumn}${lastRow}`).getValues();
    const movementValues = sheet.getRange(`${movementColumn}${startRow}:${movementColumn}${lastRow}`).getValues();

    // Merged simulation: while inactive, search for S (wMax >= Ucap(t) + Config's Slope Threshold, per firmware) once the
    // movement has stabilised (UCAP_STABILISATION samples); once active, E is checked first, R only in the "else" branch
    // (firmware only ever evaluates R when E did not fire that cycle, so E and R are mutually exclusive on any given row).
    // The block ends on our own E/R or a real X (X-triggering isn't simulated, see X.md) - a real E/R no longer force-ends
    // the block now that S itself is simulated rather than trusted.
    const simOutput: string[][] = [];
    const simValueOutput: (number | string)[][] = [];
    let active = false;
    let sIndex = -1;
    let cntStartValue = ucapTableLength;
    let baseline = 0;

    for (let i = 0; i < infoValues.length; i++) {
      const info = String(infoValues[i][0]).trim();

      if (!active) {
        const movement = movementValues[i][0] as number;
        const found = movement >= ucapStabilisation ? findSlopeBaseline(ucapFiltValues, i) : null;
        const current = ucapFiltValues[i][0] as number;
        if (found !== null && found.wMax >= current + cfg.slopeThreshold) {
          cntStartValue = found.kWinner === 4 ? ucapTableLength - 5 : ucapTableLength - found.kWinner;
          baseline = found.wMax;
          active = true;
          sIndex = i;
          simOutput.push(["S"]);
          simValueOutput.push([found.wMax]);
        } else {
          simOutput.push([""]);
          simValueOutput.push([""]);
        }
        continue;
      }

      const current = ucapFiltValues[i][0] as number;
      const previous = ucapFiltValues[i - 1][0] as number;
      const threshold = compThresholdValues[i][0] as number;
      const eDetected = current + threshold <= baseline && current < previous;

      let marker = "";
      let value: number | string = "";
      if (eDetected) {
        marker = "E";
        value = `${baseline - current} (${baseline} - ${current})`;
      } else {
        const samplesSinceS = i - sIndex;
        const cnt = Math.max(cntStartValue - samplesSinceS, 0);
        const prevIdx = i - prevSamples;
        const enoughSamples = cnt <= ucapTableLength - 11;
        const prevValue = prevIdx >= 0 ? (ucapFiltValues[prevIdx][0] as number) : null;
        const drop = prevValue !== null ? prevValue - current : null;
        const dropSmall = drop !== null && drop <= cfg.resetThreshold;
        if ((enoughSamples && dropSmall) || cnt <= 0) {
          marker = "R";
          value = drop !== null ? `${drop} (${prevValue} - ${current})` : "";
        }
      }

      simOutput.push([marker]);
      simValueOutput.push([value]);

      // Our own E/R ends the block; a real X also ends it (X-triggering isn't simulated, only its reset side-effect).
      if (marker !== "" || info === "X") {
        active = false;
      }
    }

    writeSimulationColumn(sheet, simOutputColumn, simOutputHeader, startRow, lastRow, simOutput);

    sheet.getRange(`${simValueColumn}1`).setValue(simValueHeader);
    sheet.getRange(`${simValueColumn}${startRow}:${simValueColumn}${lastRow}`).setValues(simValueOutput);

    formatGeneratedColumns(sheet, simOutputColumn, simValueColumn, startRow, lastRow);
    highlightMatchRows(sheet, [simOutputColumn, simValueColumn], startRow, lastRow, simOutput, infoValues);

    console.log(`"${cfg.sheetName}": updated columns ${simOutputColumn} (S/R/E simulation), ${simValueColumn} (threshold used), reset threshold=${cfg.resetThreshold}, rows ${startRow}-${lastRow}.`);
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

// Writes the merged simulation column.
function writeSimulationColumn(sheet: ExcelScript.Worksheet, column: string, header: string, startRow: number, lastRow: number, output: string[][]): void {
  sheet.getRange(`${column}1`).setValue(header);
  sheet.getRange(`${column}${startRow}:${column}${lastRow}`).setValues(output);
}

// Colors every row the same way across all given columns: green = simulated S/R/E matches the real Symbol (S, R or E), red = mismatch.
// Applied on top of the row banding from formatGeneratedColumns, only on rows that actually have a marker to compare.
function highlightMatchRows(sheet: ExcelScript.Worksheet, columns: string[], startRow: number, lastRow: number, output: string[][], infoValues: (string | number | boolean)[][]): void {
  for (let i = 0; i < infoValues.length; i++) {
    const info = String(infoValues[i][0]).trim();
    const realMarker = info === "S" || info === "E" || info === "R" ? info : "";
    const simMarker = output[i][0];
    if (realMarker === "" && simMarker === "") {
      continue;
    }

    const color = realMarker === simMarker ? "#C6EFCE" : "#FFC7CE";
    const row = startRow + i;
    for (const column of columns) {
      sheet.getRange(`${column}${row}`).getFormat().getFill().setColor(color);
    }
  }
}

// Bold + light-filled header; subtle alternating row shading on the data; thin grid borders; columns auto-fit to the data, then widened if needed so header words don't split mid-word.
function formatGeneratedColumns(sheet: ExcelScript.Worksheet, startColumn: string, endColumn: string, startRow: number, lastRow: number): void {
  const headerRange = sheet.getRange(`${startColumn}1:${endColumn}1`);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().setWrapText(true);
  headerRange.getFormat().getFill().setColor("#D9E1F2");

  const dataRange = sheet.getRange(`${startColumn}${startRow}:${endColumn}${lastRow}`);
  dataRange.getFormat().getFill().clear();
  applyRowBanding(sheet, startColumn, endColumn, startRow, lastRow);

  applyBorders(sheet.getRange(`${startColumn}1:${endColumn}${lastRow}`));
  sheet.getRange(`${startColumn}1:${endColumn}${lastRow}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  sheet.getRange(`${startColumn}:${endColumn}`).getFormat().autofitColumns();
  ensureHeaderColumnWidths(sheet, startColumn, endColumn);
  headerRange.getFormat().autofitRows();
}

// Very light shading on every other data row (header excluded) so consecutive rows are easier to tell apart.
function applyRowBanding(sheet: ExcelScript.Worksheet, startColumn: string, endColumn: string, startRow: number, lastRow: number): void {
  for (let row = startRow + 1; row <= lastRow; row += 2) {
    sheet.getRange(`${startColumn}${row}:${endColumn}${row}`).getFormat().getFill().setColor("#F2F2F2");
  }
}

// Widens a column if its auto-fit width would force the header's longest word to split across lines.
function ensureHeaderColumnWidths(sheet: ExcelScript.Worksheet, startColumn: string, endColumn: string): void {
  for (let code = startColumn.charCodeAt(0); code <= endColumn.charCodeAt(0); code++) {
    const column = String.fromCharCode(code);
    const header = String(sheet.getRange(`${column}1`).getValue());
    const longestWord = header.split(" ").reduce((max, word) => Math.max(max, word.length), 0);
    const minWidth = longestWord * 10 + 20; // generous per-character estimate (headers are bold), just enough to avoid mid-word wraps
    const columnRange = sheet.getRange(`${column}:${column}`);
    if (columnRange.getFormat().getColumnWidth() < minWidth) {
      columnRange.getFormat().setColumnWidth(minWidth);
    }
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

// Finds the t-1..t-4 sample with the highest Ucap Filt (as in firmware: t-4 is the default candidate, overridden only by a strictly greater t-1/t-2/t-3).
function findSlopeBaseline(ucapFiltValues: (string | number | boolean)[][], i: number): { wMax: number; kWinner: number } | null {
  let kWinner = 4;
  let wMax: number | null = i - 4 >= 0 ? (ucapFiltValues[i - 4][0] as number) : null;

  for (let k = 1; k <= 3; k++) {
    const idx = i - k;
    if (idx < 0) {
      continue;
    }
    const value = ucapFiltValues[idx][0] as number;
    if (wMax === null || value > wMax) {
      wMax = value;
      kWinner = k;
    }
  }

  return wMax === null ? null : { wMax, kWinner };
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
