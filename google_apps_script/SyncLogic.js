/**
 * Pure logic for the two telesales scripts.
 * Safe in Apps Script (globals) and Node tests (module.exports).
 *
 * Script 1: inspect / drop rows whose column L is empty.
 * Script 2: append Metabase rows under the existing sheet data.
 */

var COLUMN_L_INDEX = 11; // 0-based index for column L

function isBlankCell(value) {
  if (value === null || value === undefined) {
    return true;
  }
  if (typeof value === "number" && isNaN(value)) {
    return true;
  }
  var text = String(value).trim();
  if (!text) {
    return true;
  }
  var lower = text.toLowerCase();
  return lower === "nan" || lower === "none" || lower === "null" || lower === "#n/a";
}

function normalizeHeaderName(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value)
    .replace(/\u200f|\u200e/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function inspectEmptyColumnRows(rows, columnIndex, hasHeader) {
  if (columnIndex === undefined) {
    columnIndex = COLUMN_L_INDEX;
  }
  if (hasHeader === undefined) {
    hasHeader = true;
  }
  if (!rows || !rows.length) {
    return {
      header_rows: 0,
      data_rows: 0,
      blank_sheet_rows: [],
      blank_count: 0,
      keep_count: 0,
    };
  }
  var start = hasHeader ? 1 : 0;
  var blankSheetRows = [];
  var keepCount = 0;
  var i;
  for (i = start; i < rows.length; i++) {
    var row = rows[i] || [];
    var cell = row.length > columnIndex ? row[columnIndex] : "";
    if (isBlankCell(cell)) {
      blankSheetRows.push(i + 1);
    } else {
      keepCount += 1;
    }
  }
  return {
    header_rows: hasHeader ? 1 : 0,
    data_rows: rows.length - start,
    blank_sheet_rows: blankSheetRows,
    blank_count: blankSheetRows.length,
    keep_count: keepCount,
  };
}

function keepRowsWithFilledColumn(rows, columnIndex, hasHeader) {
  if (columnIndex === undefined) {
    columnIndex = COLUMN_L_INDEX;
  }
  if (hasHeader === undefined) {
    hasHeader = true;
  }
  if (!rows || !rows.length) {
    return [];
  }
  function hasValue(row) {
    var cell = row && row.length > columnIndex ? row[columnIndex] : "";
    return !isBlankCell(cell);
  }
  if (!hasHeader) {
    return rows.filter(hasValue).map(function (row) {
      return row.slice();
    });
  }
  var header = rows[0].slice();
  var kept = [];
  var i;
  for (i = 1; i < rows.length; i++) {
    if (hasValue(rows[i])) {
      kept.push(rows[i].slice());
    }
  }
  return [header].concat(kept);
}

function stringifyCell(value) {
  if (isBlankCell(value)) {
    return "";
  }
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return value.toISOString();
  }
  return value;
}

function alignNewRows(newBody, newHeader, sheetHeader) {
  if (!newBody || !newBody.length) {
    return [];
  }
  var header = sheetHeader ? sheetHeader.slice() : [];
  var headerKeys = header.map(normalizeHeaderName);
  var colIndexByKey = {};
  (newHeader || []).forEach(function (name, index) {
    var key = normalizeHeaderName(name);
    if (key) {
      colIndexByKey[key] = index;
    }
  });
  var overlap = headerKeys.filter(function (key) {
    return key && Object.prototype.hasOwnProperty.call(colIndexByKey, key);
  });

  return newBody.map(function (record) {
    if (overlap.length) {
      return headerKeys.map(function (key) {
        if (!Object.prototype.hasOwnProperty.call(colIndexByKey, key)) {
          return "";
        }
        return stringifyCell(record[colIndexByKey[key]]);
      });
    }
    var values = (record || []).map(stringifyCell);
    if (values.length < header.length) {
      while (values.length < header.length) {
        values.push("");
      }
    } else if (values.length > header.length) {
      values = values.slice(0, header.length);
    }
    return values;
  });
}

function appendAlignedRows(sheetRows, newHeader, newBody) {
  var aligned = alignNewRows(newBody, newHeader, sheetRows && sheetRows.length ? sheetRows[0] : newHeader);
  if (!sheetRows || !sheetRows.length) {
    return [newHeader.slice()].concat(aligned.length ? aligned : newBody.map(function (row) {
      return row.slice();
    }));
  }
  return sheetRows.map(function (row) {
    return row.slice();
  }).concat(aligned);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    COLUMN_L_INDEX: COLUMN_L_INDEX,
    isBlankCell: isBlankCell,
    inspectEmptyColumnRows: inspectEmptyColumnRows,
    keepRowsWithFilledColumn: keepRowsWithFilledColumn,
    alignNewRows: alignNewRows,
    appendAlignedRows: appendAlignedRows,
  };
}
