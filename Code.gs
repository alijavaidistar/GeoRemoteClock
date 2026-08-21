/**
 * Geometris Clock — Google Apps Script backend
 *
 * Paste this into the script editor of your Google Sheet
 * (Extensions -> Apps Script), then deploy as a Web App.
 *
 * Expected sheet tabs:
 *   Employees: PIN | Name | Active
 *   Punches:   Date | PIN | Name | Clock In | Clock Out | Hours | InEpoch
 *
 * Times come from the SERVER (Google's clock), shown in the
 * spreadsheet's timezone (File -> Settings -> Time zone).
 * InEpoch is only used to calculate hours; hide the column if you like.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // avoid two punches writing at the same moment
  try {
    var data = JSON.parse(e.postData.contents);
    var pin = String(data.pin || '').trim();
    if (!pin) return jsonOut({ ok: false, error: 'No PIN entered.' });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var employees = ss.getSheetByName('Employees');
    var punches = ss.getSheetByName('Punches');
    if (!employees || !punches) {
      return jsonOut({ ok: false, error: 'Sheet tabs "Employees" and "Punches" not found.' });
    }

    // --- Server time, in the spreadsheet's timezone ---
    var now = new Date();
    var tz = ss.getSpreadsheetTimeZone();
    var epoch = now.getTime();
    var dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var timeStr = Utilities.formatDate(now, tz, 'hh:mm a');

    // --- Look up the PIN in Employees ---
    var empRows = employees.getDataRange().getValues();
    var name = null;
    for (var i = 1; i < empRows.length; i++) {
      var rowPin = String(empRows[i][0]).trim();
      var active = empRows[i][2];
      if (rowPin === pin && active !== false && String(active).toUpperCase() !== 'FALSE') {
        name = String(empRows[i][1]);
        break;
      }
    }
    if (!name) return jsonOut({ ok: false, error: 'Invalid PIN.' });

    // --- Find this PIN's most recent punch row ---
    var pRows = punches.getDataRange().getValues();
    var lastRowIndex = -1; // 0-based index into pRows
    for (var j = pRows.length - 1; j >= 1; j--) {
      if (String(pRows[j][1]).trim() === pin) {
        lastRowIndex = j;
        break;
      }
    }

    var hasOpenSession =
      lastRowIndex !== -1 &&
      pRows[lastRowIndex][3] !== '' &&
      (pRows[lastRowIndex][4] === '' || pRows[lastRowIndex][4] === null);

    if (!hasOpenSession) {
      // --- CLOCK IN: new session row (apostrophes keep values as plain text) ---
      punches.appendRow(["'" + dateStr, "'" + pin, name, "'" + timeStr, '', '', epoch]);
      return jsonOut({ ok: true, action: 'in', name: name, time: timeStr });
    }

    // --- CLOCK OUT: close the open session ---
    var sheetRow = lastRowIndex + 1; // 1-based row number in the sheet
    var inEpoch = Number(pRows[lastRowIndex][6]);
    var hours = Math.round(((epoch - inEpoch) / 3600000) * 100) / 100;
    if (!inEpoch || hours < 0) hours = 0; // bad/missing data safety net

    punches.getRange(sheetRow, 5).setValue("'" + timeStr);
    punches.getRange(sheetRow, 6).setValue(hours).setNumberFormat('0.00');

    // --- Total of all today's sessions for this PIN (same client-side date) ---
    var totalToday = hours;
    for (var k = 1; k < pRows.length; k++) {
      if (k === lastRowIndex) continue;
      if (String(pRows[k][1]).trim() === pin &&
          String(pRows[k][0]).trim() === dateStr &&
          pRows[k][5] !== '') {
        totalToday += Number(pRows[k][5]);
      }
    }
    totalToday = Math.round(totalToday * 100) / 100;

    return jsonOut({ ok: true, action: 'out', name: name, time: timeStr, hours: hours, totalToday: totalToday });
  } catch (err) {
    return jsonOut({ ok: false, error: 'Server error: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

// Quick health check: open the web app URL in a browser and you should see OK.
function doGet() {
  return ContentService.createTextOutput('OK - Geometris Clock backend is running.');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
