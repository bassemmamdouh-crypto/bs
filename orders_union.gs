const UNION_ORDERS_CONFIG = {
  targetSheetName: "All Orders",
  headerRow: 1,
  includeSheets: [], // Example: ["Orders Jan", "Orders Feb"]
  excludeSheets: [], // Example: ["Notes", "Archive"]
  sourceNamePrefix: "", // Example: "Orders "
  addSourceSheetColumn: true,
  sourceSheetColumnName: "Source Sheet",
  clearTargetBeforeWrite: true
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Orders Tools")
    .addItem("Union Orders", "unionOrdersData")
    .addToUi();
}

function unionOrdersData() {
  const config = UNION_ORDERS_CONFIG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = getOrCreateSheet_(ss, config.targetSheetName);
  const sourceSheets = resolveSourceSheets_(ss, config);

  if (sourceSheets.length === 0) {
    throw new Error("No source sheets found. Update UNION_ORDERS_CONFIG.");
  }

  const masterHeaders = [];
  const masterHeaderSet = new Set();
  const rowObjects = [];

  sourceSheets.forEach((sheet) => {
    const values = sheet.getDataRange().getValues();
    if (values.length < config.headerRow) {
      return;
    }

    const headerRowValues = values[config.headerRow - 1];
    const normalizedHeaders = headerRowValues.map((value, index) => {
      const header = normalizeHeader_(value);
      return header || "Column " + (index + 1);
    });

    normalizedHeaders.forEach((header) => {
      if (!masterHeaderSet.has(header)) {
        masterHeaderSet.add(header);
        masterHeaders.push(header);
      }
    });

    for (let rowIndex = config.headerRow; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      if (isBlankRow_(row)) {
        continue;
      }

      const rowObject = {};
      normalizedHeaders.forEach((header, colIndex) => {
        rowObject[header] = row[colIndex];
      });

      if (config.addSourceSheetColumn) {
        rowObject[config.sourceSheetColumnName] = sheet.getName();
      }

      rowObjects.push(rowObject);
    }
  });

  if (config.addSourceSheetColumn && !masterHeaderSet.has(config.sourceSheetColumnName)) {
    masterHeaders.push(config.sourceSheetColumnName);
  }

  if (config.clearTargetBeforeWrite) {
    targetSheet.clearContents();
  }

  if (masterHeaders.length === 0) {
    targetSheet.getRange(1, 1).setValue("No data found.");
    return;
  }

  targetSheet.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders]);

  if (rowObjects.length > 0) {
    const output = rowObjects.map((rowObject) => {
      return masterHeaders.map((header) => {
        return Object.prototype.hasOwnProperty.call(rowObject, header) ? rowObject[header] : "";
      });
    });

    targetSheet
      .getRange(2, 1, output.length, masterHeaders.length)
      .setValues(output);
  }

  targetSheet.autoResizeColumns(1, masterHeaders.length);
  SpreadsheetApp.flush();
}

function resolveSourceSheets_(ss, config) {
  const allSheets = ss.getSheets();
  const includeSet = new Set(config.includeSheets || []);
  const excludeSet = new Set((config.excludeSheets || []).concat([config.targetSheetName]));

  let sourceSheets = allSheets.filter((sheet) => !excludeSet.has(sheet.getName()));

  if (includeSet.size > 0) {
    sourceSheets = sourceSheets.filter((sheet) => includeSet.has(sheet.getName()));
  }

  if (config.sourceNamePrefix) {
    sourceSheets = sourceSheets.filter((sheet) => sheet.getName().startsWith(config.sourceNamePrefix));
  }

  return sourceSheets;
}

function getOrCreateSheet_(ss, sheetName) {
  const existing = ss.getSheetByName(sheetName);
  return existing || ss.insertSheet(sheetName);
}

function normalizeHeader_(value) {
  return String(value || "").trim();
}

function isBlankRow_(row) {
  return !row.some((cell) => String(cell).trim() !== "");
}
