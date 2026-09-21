/**
 * Script 2 — pull Metabase question 590 and append it under the rows already
 * on telesales_test. Does not delete anything.
 *
 * Requires Script properties METABASE_USERNAME and METABASE_PASSWORD.
 */

function appendMetabaseData() {
  var sheet = getTargetSheet_();
  var existing = readSheetRows_(sheet);
  var metabase = fetchMetabaseQuestion_(TELESALES_CONFIG.QUESTION_ID);

  var aligned;
  var appendedCount;
  if (!existing.length) {
    aligned = [metabase.header.slice()].concat(metabase.body);
    replaceSheetRows_(sheet, aligned);
    appendedCount = metabase.body.length;
  } else {
    aligned = alignNewRows(metabase.body, metabase.header, existing[0]);
    appendSheetRows_(sheet, aligned);
    appendedCount = aligned.length;
  }

  var summary = {
    existing_rows: existing.length,
    appended_rows: appendedCount,
    question_id: TELESALES_CONFIG.QUESTION_ID,
  };
  Logger.log(JSON.stringify(summary));
  toast_(
    "Appended " + summary.appended_rows + " row(s) from Metabase question " +
      TELESALES_CONFIG.QUESTION_ID + "."
  );
  return summary;
}

function fetchMetabaseQuestion_(questionId) {
  var props = PropertiesService.getScriptProperties();
  var username = props.getProperty("METABASE_USERNAME");
  var password = props.getProperty("METABASE_PASSWORD");
  if (!username || !password) {
    throw new Error(
      "Set Script properties METABASE_USERNAME and METABASE_PASSWORD " +
        "(Project Settings in the Apps Script editor)."
    );
  }

  var sessionResponse = UrlFetchApp.fetch(TELESALES_CONFIG.METABASE_BASE_URL + "/session", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ username: username, password: password }),
    muteHttpExceptions: true,
    followRedirects: true,
  });
  if (sessionResponse.getResponseCode() < 200 || sessionResponse.getResponseCode() >= 300) {
    throw new Error(
      "Metabase session failed (" + sessionResponse.getResponseCode() + "): " +
        sessionResponse.getContentText().substring(0, 300)
    );
  }
  var sessionJson = JSON.parse(sessionResponse.getContentText());
  if (!sessionJson || !sessionJson.id) {
    throw new Error("Metabase session response is missing id.");
  }

  var csvResponse = UrlFetchApp.fetch(
    TELESALES_CONFIG.METABASE_BASE_URL + "/card/" + questionId + "/query/csv",
    {
      method: "post",
      contentType: "application/json",
      headers: { "X-Metabase-Session": sessionJson.id },
      payload: JSON.stringify({ parameters: [] }),
      muteHttpExceptions: true,
      followRedirects: true,
    }
  );
  if (csvResponse.getResponseCode() < 200 || csvResponse.getResponseCode() >= 300) {
    throw new Error(
      "Metabase CSV query failed (" + csvResponse.getResponseCode() + "): " +
        csvResponse.getContentText().substring(0, 300)
    );
  }

  var csvText = csvResponse.getContentText();
  if (csvText.charCodeAt(0) === 0xfeff) {
    csvText = csvText.substring(1);
  }
  var parsed = Utilities.parseCsv(csvText);
  if (!parsed || !parsed.length) {
    return { header: [], body: [] };
  }
  return { header: parsed[0], body: parsed.slice(1) };
}
