/**
 * Pure merge logic for the telesales daily refresh.
 * Safe to run in Apps Script (global functions) and in Node tests (module.exports).
 */

var COLUMN_P_INDEX = 15;

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

function keepRowsWithColumnP(rows, hasHeader) {
  if (hasHeader === undefined) {
    hasHeader = true;
  }
  if (!rows || !rows.length) {
    return [];
  }
  function hasValueInP(row) {
    var cell = row && row.length > COLUMN_P_INDEX ? row[COLUMN_P_INDEX] : "";
    return !isBlankCell(cell);
  }
  if (!hasHeader) {
    return rows.filter(hasValueInP).map(function (row) {
      return row.slice();
    });
  }
  var header = rows[0].slice();
  var kept = [];
  for (var i = 1; i < rows.length; i++) {
    if (hasValueInP(rows[i])) {
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

function mergeKeptAndNew(sheetRows, newHeader, newBody) {
  var kept = keepRowsWithColumnP(sheetRows, true);
  if (!kept.length) {
    return [newHeader.slice()].concat(newBody.map(function (row) {
      return row.slice();
    }));
  }
  var header = kept[0];
  return kept.concat(alignNewRows(newBody, newHeader, header));
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    COLUMN_P_INDEX: COLUMN_P_INDEX,
    isBlankCell: isBlankCell,
    keepRowsWithColumnP: keepRowsWithColumnP,
    alignNewRows: alignNewRows,
    mergeKeptAndNew: mergeKeptAndNew,
  };
}
