/**
 * Creates a custom menu in Google Sheets.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Room Pricing Tools')
      .addItem('1. Unmerge and Fill (Fast)', 'unmergeAndFillFast')
      .addItem('2. Remove Blank Rows', 'removeBlankRows')
      .addItem('3. Add Missing Dates (Fill 0s)', 'addMissingDates')
      .addItem('4. Repeat Unit Types Across Dates', 'addDatesForUnitTypes')
      .addItem('5. Check Unit Type / Date Integrity', 'verifyUnitTypeIntegrity')
      .addToUi();
}

/**
 * Tool 1: Unmerge and fill the categories.
 */
function unmergeAndFillFast() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var response1 = ui.prompt('Enter the first merged cell (e.g., B2):');
  if (response1.getSelectedButton() != ui.Button.OK) return;
  var firstCellA1 = response1.getResponseText().trim().toUpperCase();

  try {
    var firstRange = sheet.getRange(firstCellA1);
  } catch (e) {
    ui.alert("Invalid cell reference provided. Please try again.");
    return;
  }

  var colIndex = firstRange.getColumn();
  var targetColIndex = colIndex - 1;

  if (targetColIndex < 1) {
    ui.alert("The chosen cell is in Column A. There is no column in front of it to place values.");
    return;
  }

  var firstRow = firstRange.getRow();
  var lastRow = sheet.getLastRow();

  if (lastRow < firstRow) {
    ui.alert("No data found below the starting cell.");
    return;
  }

  var numRows = lastRow - firstRow + 1;
  var targetColumnRange = sheet.getRange(firstRow, colIndex, numRows, 1);

  var mergedRanges = targetColumnRange.getMergedRanges();
  for (var i = 0; i < mergedRanges.length; i++) {
    var mRange = mergedRanges[i];
    var val = mRange.getCell(1, 1).getValue();

    if (String(val).toLowerCase().indexOf("unit type:") === -1 && val !== "") {
      mRange.breakApart();
    }
  }

  var values = sheet.getRange(firstRow, colIndex, numRows, 1).getValues();
  var outputValues = [];
  var currentUnitType = "";

  for (var r = 0; r < numRows; r++) {
    var cellValue = values[r][0];

    if (cellValue !== "" && String(cellValue).toLowerCase().indexOf("unit type:") === -1) {
      if (isNaN(Date.parse(cellValue)) && isNaN(cellValue)) {
         currentUnitType = cellValue;
         outputValues.push([""]);
         continue;
      }
    }

    if (currentUnitType !== "" && cellValue !== "") {
      if (String(cellValue).toLowerCase().indexOf("unit type:") === -1) {
        outputValues.push([currentUnitType]);
      } else {
        outputValues.push([""]);
      }
    } else {
      outputValues.push([""]);
    }
  }

  sheet.getRange(firstRow, targetColIndex, numRows, 1).setValues(outputValues);
  ui.alert("Success! The sheet was processed instantly using high-speed batching.");
}

/**
 * Tool 2: Delete rows based on a target column being blank, keeping Row 1.
 */
function removeBlankRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt('Enter the column letter to check for blanks (e.g., A):');
  if (response.getSelectedButton() != ui.Button.OK) return;

  var colLetter = response.getResponseText().trim().toUpperCase();

  if (!colLetter.match(/^[A-Z]+$/)) {
    ui.alert("Please enter a valid column letter (like A, B, or C).");
    return;
  }

  try {
    var colIndex = sheet.getRange(colLetter + "1").getColumn();
  } catch (e) {
    ui.alert("Invalid column letter.");
    return;
  }

  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert("There is not enough data to process.");
    return;
  }

  var values = sheet.getRange(1, colIndex, lastRow, 1).getValues();

  var deleteStart = -1;
  var deleteCount = 0;

  // Loop backwards so deleting rows doesn't shift the sequence. Stops at r = 2
  // so Row 1 is never deleted.
  for (var r = lastRow; r >= 2; r--) {
    var cellValue = String(values[r - 1][0]).trim();

    if (cellValue === "") {
      if (deleteStart === -1) {
        deleteStart = r;
        deleteCount = 1;
      } else {
        deleteStart = r;
        deleteCount++;
      }
    } else {
      if (deleteStart !== -1) {
        sheet.deleteRows(deleteStart, deleteCount);
        deleteStart = -1;
        deleteCount = 0;
      }
    }
  }

  if (deleteStart !== -1) {
    sheet.deleteRows(deleteStart, deleteCount);
  }

  ui.alert("Success! All rows where Column " + colLetter + " was blank have been expunged. Row 1 was safely preserved.");
}

/**
 * Tool 3: Add missing dates for each room type, filling Nights/Revenue with 0.
 * Assumes: Column A = Room Type (already filled, no blanks), Column B = Date,
 * Column C = Nights, Column D = Revenue. Run this AFTER Tools 1 and 2.
 */
function addMissingDates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No data found.");
    return;
  }

  var colResp = ui.prompt('Enter the column letter that contains the dates (e.g., B):');
  if (colResp.getSelectedButton() != ui.Button.OK) return;

  var dateColLetter = colResp.getResponseText().trim().toUpperCase();
  if (!dateColLetter.match(/^[A-Z]+$/)) {
    ui.alert("Please enter a valid column letter (like A, B, or C).");
    return;
  }

  var dateColIndex;
  try {
    dateColIndex = sheet.getRange(dateColLetter + "1").getColumn();
  } catch (e) {
    ui.alert("Invalid column letter.");
    return;
  }

  var dateColValues = sheet.getRange(2, dateColIndex, lastRow - 1, 1).getValues();
  var allDates = [];
  for (var d = 0; d < dateColValues.length; d++) {
    var val = dateColValues[d][0];
    if (val === "" || val === null) continue;
    var parsed = new Date(val); // handles Date objects and "MM/DD/YY" strings
    if (!isNaN(parsed.getTime())) {
      parsed.setHours(0, 0, 0, 0);
      allDates.push(parsed.getTime());
    }
  }

  if (allDates.length === 0) {
    ui.alert("No valid dates found in column " + dateColLetter + ".");
    return;
  }

  var startDate = new Date(Math.min.apply(null, allDates));
  var endDate = new Date(Math.max.apply(null, allDates));

  if (endDate.getTime() < startDate.getTime()) {
    ui.alert("End date is before start date. Please run again with a valid range.");
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // Group rows into contiguous blocks by room type (Column A)
  var blocks = [];
  var currentType = null;
  var currentRows = [];

  for (var i = 0; i < data.length; i++) {
    var roomType = data[i][0];
    if (roomType !== currentType) {
      if (currentType !== null) {
        blocks.push({ roomType: currentType, rows: currentRows });
      }
      currentType = roomType;
      currentRows = [];
    }
    currentRows.push([data[i][1], data[i][2], data[i][3]]); // date, nights, revenue
  }
  if (currentType !== null) {
    blocks.push({ roomType: currentType, rows: currentRows });
  }

  var output = [];

  blocks.forEach(function (block) {
    var dateMap = {};

    block.rows.forEach(function (r) {
      var d = new Date(r[0]);
      d.setHours(0, 0, 0, 0);
      dateMap[d.getTime()] = [r[1], r[2]];
    });

    var cursor = new Date(startDate);
    while (cursor.getTime() <= endDate.getTime()) {
      var key = cursor.getTime();
      if (dateMap.hasOwnProperty(key)) {
        output.push([block.roomType, new Date(cursor), dateMap[key][0], dateMap[key][1]]);
      } else {
        output.push([block.roomType, new Date(cursor), 0, 0]);
      }
      cursor.setDate(cursor.getDate() + 1);
    }
  });

  sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  if (output.length > 0) {
    sheet.getRange(2, 1, output.length, 4).setValues(output);
    sheet.getRange(2, 2, output.length, 1).setNumberFormat("MM/dd/yy");
  }

  ui.alert("Success! Every room type now has a row for each date from " +
    Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM/dd/yy") + " to " +
    Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM/dd/yy") +
    ". Total rows now: " + output.length);
}

/**
 * Tool 4: Repeats every unit-type row for each date in a range.
 * Layout: A Property Guid | B Property code | C ID Ratecode | D Rate name
 *         E Unit type ID  | F Unit type     | G Date        | H:O rates
 *
 * A:F are copied as one intact slice, G gets the date, H:O are left blank.
 * Output order: unit type 1 for every date, then unit type 2, and so on.
 * The source is checked before writing and every written row is read back
 * and compared before the script reports success.
 */
function addDatesForUnitTypes() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No unit-type rows found below the header.');
    return;
  }

  // --- Build the template from unique A:F combinations ---
  var block = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var seen = {};
  var template = [];

  block.forEach(function (row) {
    var blank = row.every(function (c) { return c === '' || c === null; });
    if (blank) return;
    var key = rowKey_(row);
    if (seen[key]) return;
    seen[key] = true;
    template.push(row);
  });

  if (!template.length) {
    ui.alert('No unit-type rows found below the header.');
    return;
  }

  // --- Pre-flight: column E and column F must agree throughout the source ---
  var conflicts = findIdNameConflicts_(template);
  if (conflicts.length) {
    ui.alert('Source data problem — nothing was written',
      'Column E and column F disagree in your existing rows. Fix these first:\n\n' +
      conflicts.slice(0, 10).join('\n') +
      (conflicts.length > 10 ? '\n\n+ ' + (conflicts.length - 10) + ' more.' : ''),
      ui.ButtonSet.OK);
    return;
  }

  // --- Dates ---
  var start = askDateInput_(ui, 'Start date');
  if (!start) return;
  var end = askDateInput_(ui, 'End date');
  if (!end) return;
  if (end < start) {
    ui.alert('The end date falls before the start date.');
    return;
  }

  var dates = [];
  for (var d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    dates.push(new Date(d));
    if (dates.length > 1000) {
      ui.alert('That range is over 1000 dates. Split it into shorter blocks.');
      return;
    }
  }

  var confirm = ui.alert('Add rows?',
    template.length + ' unit types x ' + dates.length + ' dates = ' +
    (template.length * dates.length) + ' new rows.\n\n' +
    'Order: all dates for one unit type, then the next unit type.',
    ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  // --- Build rows: one unit type at a time, all its dates ---
  var rows = [];
  template.forEach(function (unit) {
    dates.forEach(function (date) {
      rows.push(unit.slice().concat([date, '', '', '', '', '', '', '', '']));
    });
  });

  var writeRow = lastRow + 1;
  var endRow = writeRow + rows.length - 1;
  if (sheet.getMaxRows() < endRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), endRow - sheet.getMaxRows());
  }

  sheet.getRange(writeRow, 7, rows.length, 1).setNumberFormat('M/d/yyyy');
  sheet.getRange(writeRow, 1, rows.length, 15).setValues(rows);
  SpreadsheetApp.flush();

  // --- Post-write verification: read it back and compare, row by row ---
  var written = sheet.getRange(writeRow, 1, rows.length, 15).getValues();
  var bad = [];

  for (var i = 0; i < rows.length; i++) {
    var problem = '';

    if (rowKey_(rows[i].slice(0, 7)) !== rowKey_(written[i].slice(0, 7))) {
      problem = 'A-G does not match what was sent';
    } else {
      for (var c = 7; c < 15; c++) {
        if (written[i][c] !== '' && written[i][c] !== null) {
          problem = 'column ' + String.fromCharCode(65 + c) + ' is not blank';
          break;
        }
      }
    }

    if (problem) bad.push('Row ' + (writeRow + i) + ': ' + problem);
    if (bad.length >= 10) break;
  }

  if (bad.length) {
    ui.alert('Verification FAILED',
      rows.length + ' rows were written but these do not match:\n\n' + bad.join('\n') +
      '\n\nDo not upload this file. Undo (Ctrl+Z) and try again.',
      ui.ButtonSet.OK);
    return;
  }

  ui.alert('Done and verified',
    'Added ' + rows.length + ' rows (rows ' + writeRow + '-' + endRow + ').\n\n' +
    template.length + ' unit types x ' + dates.length + ' dates.\n' +
    'Every row was read back and checked: A-F match the source exactly, ' +
    'column G holds the right date, H-O are blank.',
    ui.ButtonSet.OK);
}

/**
 * Tool 5: Audits the whole sheet. Run it before you upload anything.
 * Checks: E to F pairing, blank IDs or names, blank or duplicate dates,
 * and whether every unit type has the same set of dates.
 * Writes an "Integrity Check" sheet listing anything it finds.
 */
function verifyUnitTypeIntegrity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var sourceName = sheet.getName();

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data found below the header.');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var issues = [];
  var idToName = {};
  var nameToId = {};
  var units = {};
  var order = [];
  var dataRows = 0;

  data.forEach(function (row, i) {
    var rowNum = i + 2;
    var blank = row.slice(0, 6).every(function (c) { return c === '' || c === null; });
    if (blank) return;
    dataRows++;

    var id = normCell_(row[4]);
    var name = normCell_(row[5]);

    if (id === '') issues.push('Row ' + rowNum + ': Unit type ID (column E) is empty.');
    if (name === '') issues.push('Row ' + rowNum + ': Unit type (column F) is empty.');

    if (id !== '' && name !== '') {
      if (idToName[id] === undefined) {
        idToName[id] = { name: name, row: rowNum };
      } else if (idToName[id].name !== name) {
        issues.push('Row ' + rowNum + ': ID ' + id + ' is "' + name +
          '" here but "' + idToName[id].name + '" on row ' + idToName[id].row + '.');
      }

      if (nameToId[name] === undefined) {
        nameToId[name] = { id: id, row: rowNum };
      } else if (nameToId[name].id !== id) {
        issues.push('Row ' + rowNum + ': "' + name + '" has ID ' + id +
          ' here but ID ' + nameToId[name].id + ' on row ' + nameToId[name].row + '.');
      }
    }

    var key = rowKey_(row.slice(0, 6));
    if (!units[key]) {
      units[key] = { label: (name || '(no name)') + ' / ID ' + (id || '(none)'), dates: {}, count: 0 };
      order.push(key);
    }

    var unit = units[key];
    unit.count++;

    var dateVal = normCell_(row[6]);
    if (dateVal === '') {
      issues.push('Row ' + rowNum + ': date (column G) is empty.');
      return;
    }
    if (unit.dates[dateVal]) {
      issues.push('Row ' + rowNum + ': ' + unit.label + ' already has ' + dateVal +
        ' on row ' + unit.dates[dateVal] + '.');
    } else {
      unit.dates[dateVal] = rowNum;
    }
  });

  // Compare each unit's date set against the most common one
  var sigCount = {};
  var sigOf = {};

  order.forEach(function (key) {
    var ds = Object.keys(units[key].dates).sort();
    sigOf[key] = ds;
    var sig = ds.join(',');
    sigCount[sig] = (sigCount[sig] || 0) + 1;
  });

  var modalSig = '';
  var modalN = -1;
  Object.keys(sigCount).forEach(function (sig) {
    if (sigCount[sig] > modalN) { modalN = sigCount[sig]; modalSig = sig; }
  });
  var modalDates = modalSig === '' ? [] : modalSig.split(',');

  order.forEach(function (key) {
    var ds = sigOf[key];
    if (ds.join(',') === modalSig) return;

    var have = {};
    ds.forEach(function (x) { have[x] = true; });
    var missing = modalDates.filter(function (x) { return !have[x]; });
    var extra = ds.filter(function (x) { return modalDates.indexOf(x) === -1; });

    var msg = units[key].label + ' has ' + ds.length + ' dates instead of ' + modalDates.length;
    if (missing.length) {
      msg += '; missing ' + missing.slice(0, 5).join(', ') +
        (missing.length > 5 ? ' + ' + (missing.length - 5) + ' more' : '');
    }
    if (extra.length) {
      msg += '; unexpected ' + extra.slice(0, 5).join(', ') +
        (extra.length > 5 ? ' + ' + (extra.length - 5) + ' more' : '');
    }
    issues.push(msg + '.');
  });

  var summary = 'Sheet: ' + sourceName + '\n' +
    'Data rows: ' + dataRows + '\n' +
    'Unit type combinations: ' + order.length + '\n' +
    'Dates per unit type: ' + modalDates.length + ' (most common)\n' +
    'Expected total: ' + (order.length * modalDates.length);

  if (!issues.length) {
    ui.alert('All clear', summary + '\n\nNo problems found. Column E matches column F ' +
      'everywhere, no blanks, no duplicate dates, and every unit type has the same date set.',
      ui.ButtonSet.OK);
    return;
  }

  var report = ss.getSheetByName('Integrity Check');
  if (report) ss.deleteSheet(report);
  report = ss.insertSheet('Integrity Check');

  var out = [['Integrity check for "' + sourceName + '"']];
  summary.split('\n').forEach(function (line) { out.push([line]); });
  out.push(['']);
  out.push([issues.length + ' issue(s) found:']);
  issues.forEach(function (issue) { out.push([issue]); });

  report.getRange(1, 1, out.length, 1).setValues(out);
  report.getRange(1, 1).setFontWeight('bold');
  report.getRange(out.length - issues.length, 1).setFontWeight('bold');
  report.setColumnWidth(1, 720);

  ui.alert('Problems found',
    summary + '\n\n' + issues.length + ' issue(s). Full list is on the new ' +
    '"Integrity Check" sheet.\n\nFirst few:\n\n' + issues.slice(0, 5).join('\n'),
    ui.ButtonSet.OK);
}

/**
 * Turns a row into a comparable string. Dates become yyyy-MM-dd so a Date
 * object and the same date read back from the sheet compare as equal.
 */
function rowKey_(row) {
  return row.map(normCell_).join('\u0001');
}

function normCell_(cell) {
  if (cell === null || cell === undefined) return '';
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(cell).trim();
}

/**
 * Reports any place where one Unit type ID maps to two names, or one name
 * maps to two IDs.
 */
function findIdNameConflicts_(template) {
  var idToName = {};
  var nameToId = {};
  var found = {};

  template.forEach(function (row) {
    var id = normCell_(row[4]);
    var name = normCell_(row[5]);
    if (id === '' || name === '') return;

    if (idToName[id] === undefined) idToName[id] = name;
    else if (idToName[id] !== name) {
      found['ID ' + id + ' is used for both "' + idToName[id] + '" and "' + name + '"'] = true;
    }

    if (nameToId[name] === undefined) nameToId[name] = id;
    else if (nameToId[name] !== id) {
      found['"' + name + '" is used with both ID ' + nameToId[name] + ' and ID ' + id] = true;
    }
  });

  return Object.keys(found);
}

/**
 * Accepts 7/10/2026 or 2026-07-10. Noon keeps timezones from shifting the day.
 */
function askDateInput_(ui, label) {
  var res = ui.prompt(label, 'Format: M/D/YYYY  (e.g. 7/10/2026)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  var t = res.getResponseText().trim();

  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (m) return new Date(+m[3], +m[1] - 1, +m[2], 12);

  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], 12);

  ui.alert('Could not read "' + t + '". Use M/D/YYYY, like 7/10/2026.');
  return null;
}
