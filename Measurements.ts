function main(workbook: ExcelScript.Workbook) {
  const configSheetName = "Config";
  const rawLineColumn = "A"; // raw input is hex text: Line, Ucap, Ucap Comp, Ucap Filt, Symbol, Threshold Comp = A-F
  const rawSymbolColumn = "E"; // end of the raw range when Config's "data with threshold" = 0 (no raw Threshold Comp)
  const rawThresholdCompColumn = "F"; // end of the raw hex range when Config's "data with threshold" = 1 (A:F)

  const plineColumn = "H"; // decoded (decimal) columns, written by this script: Line, Ucap, Ucap Comp, Ucap Filt, Symbol, Threshold Comp = H-M
  const ucapFiltColumn = "K";
  const detectionColumn = "L";
  const decThresholdCompColumn = "M"; // end of the decoded range (H:M)
  const decodedHeaders = ["Line", "Ucap", "Ucap Comp", "Ucap Filt", "Symbol", "Threshold Comp"];

  const dataStartRow = 2; // row 1 holds the input column labels, real data starts at row 2
  const diffThresholdMovement = 25; // threshold for Ucap diff, based on the "Movement" counter (old nb > 25)
  const cntStart = 150;
  const sensitivityMax = 32768; // firmware: SENSITIVITY_MAX, forces the threshold for the first couple of samples after S
  const criticalDeltaPowerLine10V = 24; // firmware: CRITICAL_DELTA_POWER_LINE_10V
  const mainsWindowLength = 150; // firmware: awLineValueTable is UCAP_TABLE_LENGTH long

  const movementColumn = "O";
  const movementHeader = "Mvt";
  const numberColumn = "P";
  const numberHeader = "Cnt";
  const ucapDiffColumn = "Q";
  const ucapDiffHeader = "Ucap diff";
  const slopeColumn = "R";
  const slopeHeader = "Slope";
  const compThresholdColumn = "S";
  const compThresholdHeader = "Threshold Comp";

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

    // Step 1: raw A-F is hex text; decode into H-M as decimal (Symbol, column E/L, is copied as-is, not hex).
    // When Config's "data with threshold" = 0, the raw sheet only has A-E (no real Threshold Comp); M is left blank.
    const rawEndColumn = cfg.dataWithThreshold ? rawThresholdCompColumn : rawSymbolColumn;
    const rawValues = sheet.getRange(`${rawLineColumn}${startRow}:${rawEndColumn}${lastRow}`).getValues();
    const decodedValues: (number | string)[][] = rawValues.map((row: (string | number | boolean)[]) => [
      hexToDec(row[0]),
      hexToDec(row[1]),
      hexToDec(row[2]),
      hexToDec(row[3]),
      String(row[4]).trim(),
      cfg.dataWithThreshold ? hexToDec(row[5]) : ""
    ]);

    sheet.getRange(`${plineColumn}1:${decThresholdCompColumn}1`).setValues([decodedHeaders]);
    sheet.getRange(`${plineColumn}${startRow}:${decThresholdCompColumn}${lastRow}`).setValues(decodedValues);
    formatGeneratedColumns(sheet, plineColumn, decThresholdCompColumn, startRow, lastRow);

    // Step 2: all calculations below read from the decoded columns (Line=H, Ucap Filt=K, Symbol=L).
    const ucapFiltValues = decodedValues.map(row => [row[3]]);
    const plineValues = decodedValues.map(row => [row[0]]);
    const detectionValues = decodedValues.map(row => [row[4]]);
    const rowCount = ucapFiltValues.length;

    const movementOutput: (number | string)[][] = [];
    const numberOutput: (number | string)[][] = [];
    const ucapDiffOutput: (number | string)[][] = [];
    const slopeOutput: (number | string)[][] = [];
    const compThresholdOutput: (number | string)[][] = [];

    // Movement: counting starts once 3 consecutive zero samples in Ucap Filt are detected (counted as 1,2,3), ends on Detection "E".
    let movementActive = false;
    let movementCounter = 0;
    for (let i = 0; i < rowCount; i++) {
      movementOutput.push([0]);
    }
    for (let i = 0; i < rowCount; i++) {
      const detection = String(detectionValues[i][0]).trim();

      if (!movementActive) {
        if (i >= 2
          && (ucapFiltValues[i][0] as number) === 0
          && (ucapFiltValues[i - 1][0] as number) === 0
          && (ucapFiltValues[i - 2][0] as number) === 0) {
          movementActive = true;
          movementCounter = 3;
          movementOutput[i - 2] = [1];
          movementOutput[i - 1] = [2];
          movementOutput[i] = [3];
        }
      } else {
        movementCounter++;
        movementOutput[i] = [movementCounter];
      }

      if (movementActive && detection === "E") {
        movementActive = false;
        movementCounter = 0;
      }
    }

    // Number: cnt counts down from the row with S, to 0 or until R/E; outside such a block it is always 150.
    // Start value depends on which of t-1..t-4 had the highest Ucap Filt (as in firmware): t-4 by default -> 145, t-k (k=1..3) strictly greater -> 150-k.
    let cntActive = false;
    let sIndex = -1;
    let cntStartValue = cntStart;
    for (let i = 0; i < rowCount; i++) {
      const detection = String(detectionValues[i][0]).trim();

      if (!cntActive) {
        numberOutput.push([cntStart]);
        if (detection === "S") {
          const found = findSlopeBaseline(ucapFiltValues, i);
          const kWinner = found ? found.kWinner : 4;
          cntStartValue = kWinner === 4 ? cntStart - 5 : cntStart - kWinner;
          cntActive = true;
          sIndex = i;
        }
        continue;
      }

      const samplesSinceS = i - sIndex;
      const cnt = Math.max(cntStartValue - samplesSinceS, 0);
      numberOutput.push([cnt]);
      // X also resets the state in firmware (same as R) - we don't apply the X threshold here, just end the block.
      if (detection === "E" || detection === "R" || detection === "X" || cnt === 0) {
        cntActive = false;
      }
    }

    // Ucap diff: independent of Number - based on the Movement counter > threshold.
    for (let i = 0; i < rowCount; i++) {
      const movement = movementOutput[i][0] as number;
      if (i > 0 && movement > diffThresholdMovement) {
        const current = ucapFiltValues[i][0] as number;
        const previous = ucapFiltValues[i - 1][0] as number;
        ucapDiffOutput.push([previous - current]);
      } else {
        ucapDiffOutput.push([""]);
      }
    }

    // Slope measurement: Detection S -> baseline from max of t-1..t-4 (as in firmware), until R/E inclusive.
    // Compensated threshold: firmware's wEldSensivityThreshold for the same active block (forced to SENSITIVITY_MAX for the first samples after S).
    let baseline: number | null = null;
    let thresholdCntStart = cntStart;
    let thresholdSIndex = -1;
    for (let i = 0; i < rowCount; i++) {
      const detection = String(detectionValues[i][0]).trim();
      const currentUcapFilt = ucapFiltValues[i][0] as number;

      if (baseline === null) {
        if (detection === "S") {
          const found = findSlopeBaseline(ucapFiltValues, i);
          const kWinner = found ? found.kWinner : 4;
          baseline = found ? found.wMax : null;
          thresholdCntStart = kWinner === 4 ? cntStart - 5 : cntStart - kWinner;
          thresholdSIndex = i;
        }
        slopeOutput.push([""]);
        compThresholdOutput.push([""]);
      } else {
        slopeOutput.push([baseline - currentUcapFilt]);

        const samplesSinceS = i - thresholdSIndex;
        const forceMax = samplesSinceS < thresholdCntStart - (cntStart - 3);
        const wSensivity = testIfMainsWasOk(plineValues, i, mainsWindowLength, criticalDeltaPowerLine10V)
          ? cfg.sensitivity
          : cfg.l0Sensitivity;
        const compThreshold = forceMax
          ? sensitivityMax
          : compensateThreshold(wSensivity, plineValues[i][0] as number);
        compThresholdOutput.push([compThreshold]);

        // X also resets the state in firmware (same as R).
        if (detection === "R" || detection === "E" || detection === "X") {
          baseline = null;
        }
      }
    }

    // Row 1 already holds the input column labels; write our own headers there too, data stays at row 2+.
    sheet.getRange(`${movementColumn}1`).setValue(movementHeader);
    sheet.getRange(`${numberColumn}1`).setValue(numberHeader);
    sheet.getRange(`${ucapDiffColumn}1`).setValue(ucapDiffHeader);
    sheet.getRange(`${slopeColumn}1`).setValue(slopeHeader);
    sheet.getRange(`${compThresholdColumn}1`).setValue(compThresholdHeader);

    sheet.getRange(`${movementColumn}${startRow}:${movementColumn}${lastRow}`).setValues(movementOutput);
    sheet.getRange(`${numberColumn}${startRow}:${numberColumn}${lastRow}`).setValues(numberOutput);
    sheet.getRange(`${ucapDiffColumn}${startRow}:${ucapDiffColumn}${lastRow}`).setValues(ucapDiffOutput);
    sheet.getRange(`${slopeColumn}${startRow}:${slopeColumn}${lastRow}`).setValues(slopeOutput);
    sheet.getRange(`${compThresholdColumn}${startRow}:${compThresholdColumn}${lastRow}`).setValues(compThresholdOutput);

    formatGeneratedColumns(sheet, movementColumn, compThresholdColumn, startRow, lastRow);

    console.log(`"${cfg.sheetName}": updated columns ${movementColumn} (${movementHeader}), ${numberColumn} (${numberHeader}), ${ucapDiffColumn} (${ucapDiffHeader}), ${slopeColumn} (${slopeHeader}), ${compThresholdColumn} (${compThresholdHeader}), rows ${startRow}-${lastRow}.`);
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

// Decodes a hex string (e.g. "5E1B") into its decimal value; blank cells decode to 0.
function hexToDec(value: string | number | boolean): number {
  const str = String(value).trim();
  return str === "" ? 0 : parseInt(str, 16);
}

// Bold + word-wrapped, light-filled header; subtle alternating row shading on the data; thin grid borders; columns auto-fit to the data, then widened if needed so header words don't split mid-word.
function formatGeneratedColumns(sheet: ExcelScript.Worksheet, startColumn: string, endColumn: string, startRow: number, lastRow: number): void {
  const headerRange = sheet.getRange(`${startColumn}1:${endColumn}1`);
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().setWrapText(true);
  headerRange.getFormat().getFill().setColor("#D9E1F2");

  const dataRange = sheet.getRange(`${startColumn}${startRow}:${endColumn}${lastRow}`);
  dataRange.getFormat().getFill().clear();
  applyRowBanding(sheet, startColumn, endColumn, startRow, lastRow);

  applyBorders(sheet.getRange(`${startColumn}1:${endColumn}${lastRow}`));

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

// firmware: TorqueMeasure_PRV_CompensateThreshold - scales the threshold up as the mains line value (wLineValue) drops.
function compensateThreshold(wThreshold: number, wLineValue: number): number {
  if (wThreshold >= 32768) {
    return wThreshold;
  }

  const wLineValueForCalculations = Math.floor((200 * wLineValue) / 673);
  if (wLineValueForCalculations >= 255) {
    return wThreshold;
  }

  let lTempComp = 255 - wLineValueForCalculations;
  if (lTempComp > 60) {
    lTempComp = 60;
  }
  lTempComp = Math.floor((lTempComp * wThreshold) / 81);
  return wThreshold + lTempComp;
}

// firmware: TorqueMeasure_PRV_TestIfMainsWasOk - mains considered unstable if Pline varied by more than the critical delta over the last windowLength samples.
function testIfMainsWasOk(plineValues: (string | number | boolean)[][], i: number, windowLength: number, criticalDelta: number): boolean {
  let wMin: number | null = null;
  let wMax = 0;
  const startIdx = Math.max(0, i - (windowLength - 1));

  for (let k = startIdx; k <= i; k++) {
    const value = plineValues[k][0] as number;
    if (value !== 0 && (wMin === null || value < wMin)) {
      wMin = value;
    }
    if (value > wMax) {
      wMax = value;
    }
  }

  if (wMin === null) {
    wMin = wMax;
  }
  return wMax - wMin <= criticalDelta;
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
