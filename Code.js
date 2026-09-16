function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gmail Label Manager')
    .addItem('Export Gmail Labels', 'exportGmailLabels')
    .addItem('Create Label Plan', 'createLabelPlan')
    .addItem('Validate Label Plan', 'validateLabelPlan')
    .addItem('Preflight Label Plan', 'preflightLabelPlan')
    .addSeparator()
    .addItem('Clear Label Inventory', 'clearLabelInventory')
    .addToUi();
}


function exportGmailLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Gmail Labels';

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Clear data only. Preserve user formatting and column widths.
  sheet.clearContents();

  const headers = [
    'Full Label Path',
    'Parent Path',
    'Label Name',
    'Depth',
    'Message Count',
    'Unread Count'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const labels = GmailApp.getUserLabels();

  const rows = labels.map(label => {
    const fullName = label.getName();
    const parts = fullName.split('/');
    const labelName = parts[parts.length - 1];

    const parentPath =
      parts.length > 1
        ? parts.slice(0, -1).join('/')
        : '';

    const depth = parts.length;

    return [
      fullName,
      parentPath,
      labelName,
      depth,
      label.getThreads(0, 500).length,
      label.getUnreadCount()
    ];
  });

  rows.sort((a, b) =>
    a[0].localeCompare(
      b[0],
      undefined,
      {
        sensitivity: 'base',
        numeric: true
      }
    )
  );

  if (rows.length > 0) {
    sheet
      .getRange(2, 1, rows.length, headers.length)
      .setValues(rows.map(row => row.map(escapeSheetApostrophe_)));
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    'Export complete. ' +
    rows.length +
    ' user-created Gmail labels were found.'
  );
}


function escapeSheetApostrophe_(value) {
  // Escape only at the write boundary; getValues() returns the literal label text.
  return typeof value === 'string' && value.startsWith("'") ? "'" + value : value;
}


function clearLabelInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Gmail Labels');

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'There is no Gmail Labels sheet to clear.'
    );
    return;
  }

  sheet.clearContents();

  SpreadsheetApp.getUi().alert(
    'Label inventory cleared.'
  );
}


function createLabelPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName('Gmail Labels');

  if (!source) {
    SpreadsheetApp.getUi().alert(
      'Please run Export Gmail Labels first.'
    );
    return;
  }

  const lastRow = source.getLastRow();
  const data = lastRow < 2
    ? []
    : source.getRange(2, 1, lastRow - 1, 3).getValues()
      .filter(row => String(row[0]).trim() !== '');

  if (data.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Gmail labels were found in the Gmail Labels sheet.'
    );
    return;
  }

  const planName = 'Label Plan';
  let plan = ss.getSheetByName(planName);
  const isNewPlan = !plan;

  if (plan && plan.getLastRow() > 1) {
    const existing = plan.getRange(2, 1, plan.getLastRow() - 1, 10);
    const hasWork = existing.getValues().some(row => row.some(value => value !== '')) ||
      existing.getFormulas().some(row => row.some(formula => formula !== ''));

    if (hasWork) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Regenerate Label Plan?',
        'This will replace all planning rows in columns A:J, including proposals, ' +
        'validation, results, and notes. This cannot be undone by this tool. ' +
        'Content outside A:J will be preserved. Continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return;
      }
    }
  }

  if (!plan) {
    plan = ss.insertSheet(planName);
  }

  if (plan.getMaxRows() < data.length + 1) {
    plan.insertRowsAfter(plan.getMaxRows(), data.length + 1 - plan.getMaxRows());
  }

  // A:J is tool-owned; clear contents without changing ordinary formatting.
  plan.getRange(1, 1, plan.getMaxRows(), 10).clearContent();

  const headers = [
    'Current Full Label Path',
    'Current Parent',
    'Current Label Name',
    'Proposed Parent',
    'Proposed Label Name',
    'Proposed Full Label Path',
    'Action',
    'Validation',
    'Result',
    'Notes'
  ];

  plan.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = data.map(row => [
    row[0], // Current Full Label Path
    row[1], // Current Parent
    row[2], // Current Label Name
    row[1], // Proposed Parent
    row[2], // Proposed Label Name
    '',     // Proposed Full Label Path
    'Keep', // Action
    '',     // Validation
    '',     // Result
    ''      // Notes
  ]);

  /*
   * Clear old dropdown validation before writing the new values.
   * This prevents an obsolete validation rule from rejecting "Keep".
   */
  const maxRows = plan.getMaxRows();

  if (maxRows > 1) {
    plan
      .getRange(2, 7, maxRows - 1, 1)
      .clearDataValidations();
  }

  plan
    .getRange(2, 1, rows.length, headers.length)
    .setValues(rows.map(row => row.map(escapeSheetApostrophe_)));

  /*
   * Proposed Full Label Path
   */
  for (let r = 2; r <= rows.length + 1; r++) {
    plan.getRange(r, 6).setFormula(
      `=IF(E${r}="","",IF(D${r}="",E${r},D${r}&"/"&E${r}))`
    );
  }

  /*
   * Action dropdown
   */
  const actionRule = labelPlanActionRule_();

  plan
    .getRange(2, 7, rows.length, 1)
    .setDataValidation(actionRule);

  updateLabelPlanFormatting_(plan, rows.length);

  if (isNewPlan) {
    plan.setFrozenRows(1);
    plan.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  SpreadsheetApp.getUi().alert(
    'Label Plan created with ' + rows.length + ' existing Gmail labels.\n\n' +
    'No Gmail changes were made.'
  );
}


function updateLabelPlanFormatting_(plan, rowCount) {

  /*
   * Conditional formatting managed by the script.
   * Ordinary user formatting and column widths remain untouched.
   */
  const dataRange = plan.getRange(
    2,
    1,
    rowCount,
    10
  );

  const rules = [
    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="Ready"')
      .setBackground('#e6f4ea')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="No change"')
      .setBackground('#f8f9fa')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=LEFT($H2,5)="Error"')
      .setBackground('#fce8e6')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=LEFT($H2,6)="Review"')
      .setBackground('#fff4e5')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp
      .newConditionalFormatRule()
      .whenFormulaSatisfied('=LEFT($H2,9)="Adjusted:"')
      .setBackground('#fff4e5')
      .setRanges([dataRange])
      .build()
  ];

  const userRules = plan.getConditionalFormatRules()
    .filter(rule => !isLabelPlanStatusRule_(rule));
  plan.setConditionalFormatRules(userRules.concat(rules));

}


function isLabelPlanStatusRule_(rule) {
  const condition = rule.getBooleanCondition();
  if (!condition ||
      condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
    return false;
  }

  // Match the legacy rules narrowly so unrelated user rules survive regeneration.
  const backgrounds = {
    '=$H2="Ready"': '#e6f4ea',
    '=$H2="No change"': '#f8f9fa',
    '=LEFT($H2,5)="Error"': '#fce8e6',
    '=LEFT($H2,6)="Review"': '#fff4e5',
    '=LEFT($H2,9)="Adjusted:"': '#fff4e5'
  };
  const formula = String(condition.getCriteriaValues()[0]);
  const ranges = rule.getRanges();

  return Object.prototype.hasOwnProperty.call(backgrounds, formula) &&
    condition.getBackgroundObject()?.asRgbColor().asHexString() === backgrounds[formula] &&
    ranges.length === 1 &&
    ranges[0].getSheet().getName() === 'Label Plan' &&
    ranges[0].getRow() === 2 &&
    ranges[0].getColumn() === 1 &&
    ranges[0].getNumColumns() === 10;
}


function labelPlanActionRule_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(['Keep', 'Rename/Move', 'Create', 'Delete'], true)
    .setAllowInvalid(false)
    .build();
}


function prepareLabelPlanRows_(plan = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Label Plan')) {
  if (!plan || plan.getLastRow() < 2) return;
  const range = plan.getRange(2, 4, plan.getLastRow() - 1, 4);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const rule = labelPlanActionRule_();
  values.forEach((row, index) => {
    if (!row.some(value => value !== '')) return;
    const r = index + 2;
    if (row[2] === '' && !formulas[index][2]) {
      plan.getRange(r, 6).setFormula(
        `=IF(E${r}="","",IF(D${r}="",E${r},D${r}&"/"&E${r}))`
      );
    }
    plan.getRange(r, 7).setDataValidation(rule);
  });
  SpreadsheetApp.flush();
}


function normalizeLabelPlanActions_(plan = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Label Plan')) {
  const adjustments = new Map();
  if (!plan || plan.getLastRow() < 2) return adjustments;
  const data = plan.getRange(2, 1, plan.getLastRow() - 1, 7).getValues();
  data.forEach((row, index) => {
    const current = String(row[0] || '');
    const proposed = String(row[5] || '');
    const action = String(row[6] || '').trim();
    // An incomplete proposal is a validation error, never an inferred deletion.
    if (action === 'Delete' || !proposed.trim()) return;
    const expected = !current ? 'Create' : current === proposed ? 'Keep' : 'Rename/Move';
    if (action === expected) return;
    const reason = !current ? 'new label' : current === proposed ? 'path unchanged' : 'path changed';
    const status = 'Adjusted: ' + reason + '; action updated to ' + expected;
    plan.getRange(index + 2, 7).setValues([[expected]]);
    plan.getRange(index + 2, 8).setValues([[status]]);
    plan.getRange(index + 2, 9).clearContent();
    adjustments.set(index + 2, status);
  });
  return adjustments;
}


function onEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Label Plan' ||
      range.getLastRow() < 2 ||
      range.getColumn() > 7 || range.getLastColumn() < 4) {
    return;
  }

  const firstRow = Math.max(2, range.getRow());
  sheet.getRange(firstRow, 8, range.getLastRow() - firstRow + 1, 3)
    .clearContent();
}


function validateLabelPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plan = ss.getSheetByName('Label Plan');

  if (!plan) {
    SpreadsheetApp.getUi().alert(
      'Please create the Label Plan first.'
    );
    return;
  }

  const lastRow = plan.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert(
      'There are no rows to validate.'
    );
    return;
  }

  prepareLabelPlanRows_(plan);
  const adjustments = normalizeLabelPlanActions_(plan);
  const rowCount = lastRow - 1;
  updateLabelPlanFormatting_(plan, rowCount);

  /*
   * Columns A:J
   */
  const data = plan
    .getRange(2, 1, rowCount, 10)
    .getValues();

  const currentPathRows = {};
  const proposedPathRows = {};

  /*
   * Build indexes of current and proposed paths.
   */
  data.forEach((row, index) => {
    const sheetRow = index + 2;

    const currentPath =
      String(row[0] || '').trim();

    const proposedPath =
      String(row[5] || '').trim();

    const action =
      String(row[6] || '').trim();

    if (currentPath) {
      if (!currentPathRows[currentPath]) {
        currentPathRows[currentPath] = [];
      }

      currentPathRows[currentPath].push(sheetRow);
    }

    if (
      proposedPath &&
      action !== 'Delete'
    ) {
      if (!proposedPathRows[proposedPath]) {
        proposedPathRows[proposedPath] = [];
      }

      proposedPathRows[proposedPath].push(sheetRow);
    }
  });

  const results = data.map((row, index) => {
    if (!row.slice(0, 7).some(value => value !== '')) return [''];
    if (adjustments.has(index + 2)) return [adjustments.get(index + 2)];
    const currentPath =
      String(row[0] || '').trim();

    const proposedName =
      String(row[4] || '').trim();

    const proposedPath =
      String(row[5] || '').trim();

    const action =
      String(row[6] || '').trim();

    let validation = '';

    if (!action) {
      validation =
        'Error: action is blank';
    }

    else if (
      ![
        'Keep',
        'Rename/Move',
        'Create',
        'Delete'
      ].includes(action)
    ) {
      validation =
        'Error: invalid action';
    }

    else if (action === 'Keep') {
      if (!currentPath) {
        validation =
          'Error: no current label';
      }

      else if (!proposedPath) {
        validation =
          'Error: proposed path is blank';
      }

      else if (
        currentPath !== proposedPath
      ) {
        validation =
          'Error: path changed but action is Keep';
      }

      else {
        validation =
          'No change';
      }
    }

    else if (action === 'Rename/Move') {
      if (!currentPath) {
        validation =
          'Error: current label is blank';
      }

      else if (
        !proposedName ||
        !proposedPath
      ) {
        validation =
          'Error: proposed label is incomplete';
      }

      else if (
        currentPath === proposedPath
      ) {
        validation =
          'Review: no change detected';
      }

      else if (
        proposedPathRows[proposedPath] &&
        proposedPathRows[proposedPath].length > 1
      ) {
        validation =
          'Error: duplicate proposed path';
      }

      else if (
        currentPathRows[proposedPath] &&
        proposedPath !== currentPath
      ) {
        validation =
          'Error: destination already exists';
      }

      else {
        validation =
          'Ready';
      }
    }

    else if (action === 'Create') {
      if (currentPath) {
        validation =
          'Review: Create normally requires a blank current path';
      }

      else if (
        !proposedName ||
        !proposedPath
      ) {
        validation =
          'Error: proposed label is incomplete';
      }

      else if (
        proposedPathRows[proposedPath] &&
        proposedPathRows[proposedPath].length > 1
      ) {
        validation =
          'Error: duplicate proposed path';
      }

      else if (
        currentPathRows[proposedPath]
      ) {
        validation =
          'Error: label already exists';
      }

      else {
        validation =
          'Ready';
      }
    }

    else if (action === 'Delete') {
      if (!currentPath) {
        validation =
          'Error: current label is blank';
      }

      else {
        validation =
          'Review: deletion requested';
      }
    }

    return [validation];
  });

  /*
   * Write validation results to column H.
   */
  plan
    .getRange(2, 8, results.length, 1)
    .setValues(results);

  const validationValues =
    results.map(row => row[0]);

  const errors =
    validationValues.filter(
      value => value.startsWith('Error')
    ).length;

  const reviews =
    validationValues.filter(
      value => value.startsWith('Review')
    ).length;

  const ready =
    validationValues.filter(
      value => value === 'Ready'
    ).length;

  const noChange =
    validationValues.filter(
      value => value === 'No change'
    ).length;

  SpreadsheetApp.getUi().alert(
    'Validation complete.\n\n' +
    'Ready: ' + ready + '\n' +
    'No change: ' + noChange + '\n' +
    'Review: ' + reviews + '\n' +
    'Errors: ' + errors + '\n' +
    'Adjusted: ' + adjustments.size +
    (adjustments.size > 0
      ? '\n\nReview adjusted rows and run Validate Label Plan again.' : '')
  );
}


function preflightLabelPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const plan = ss.getSheetByName('Label Plan');
  if (!plan) {
    ui.alert('Please create the Label Plan first.');
    return;
  }

  const lastRow = plan.getLastRow();
  if (lastRow < 2) {
    ui.alert('There are no rows to preflight.');
    return;
  }

  prepareLabelPlanRows_(plan);
  const data = plan.getRange(2, 1, lastRow - 1, 10).getValues();
  const resultRange = plan.getRange(2, 9, data.length, 1);
  resultRange.clearContent();
  const rows = data.map(row => ({
    active: row.slice(0, 7).some(value => value !== ''),
    current: String(row[0] || ''),
    parent: String(row[3] || ''),
    name: String(row[4] || ''),
    destination: String(row[5] || ''),
    action: String(row[6] || '').trim(),
    validation: String(row[7] || '').trim()
  }));

  // Validation is a prerequisite, but every run still checks a fresh Gmail snapshot.
  let livePaths;
  try {
    livePaths = new Set(GmailApp.getUserLabels().map(label => label.getName()));
  } catch (error) {
    resultRange.setValues(rows.map(row => [row.active
      ? 'Blocked: live Gmail labels could not be read' : '']));
    ui.alert('Preflight could not read live Gmail labels. No Gmail changes were made.\n\n' +
      'Ready operations: 0\nDeletion reviews: 0\nBlocked operations: ' +
      rows.filter(row => row.active).length + '\nUnchanged rows: 0');
    return;
  }

  const plannedSources = new Set(rows.filter(row => row.current).map(row => row.current));
  const missing = [...plannedSources].filter(path => !livePaths.has(path));
  const added = [...livePaths].filter(path => !plannedSources.has(path));
  const destinations = new Map();
  const sources = new Map();
  rows.forEach(row => {
    if (row.current) {
      sources.set(row.current, (sources.get(row.current) || 0) + 1);
    }
    if (['Keep', 'Rename/Move', 'Create'].includes(row.action) && row.destination) {
      destinations.set(row.destination, (destinations.get(row.destination) || 0) + 1);
    }
  });

  const counts = { ready: 0, deletion: 0, blocked: 0, unchanged: 0 };
  const results = rows.map(row => {
    if (!row.active) {
      return [''];
    }

    let problem = '';
    const reconstructed = !row.name ? '' : row.parent ? row.parent + '/' + row.name : row.name;
    if (!['Keep', 'Rename/Move', 'Create', 'Delete'].includes(row.action)) {
      problem = 'invalid or blank action';
    } else if (row.action === 'Create' && row.current) {
      problem = 'Create requires a blank current label path';
    } else if (row.action !== 'Create' && !row.current) {
      problem = 'source label is blank';
    } else if (row.action !== 'Create' && !livePaths.has(row.current)) {
      problem = 'stale plan: source label no longer exists in Gmail';
    } else if (row.current && sources.get(row.current) > 1) {
      problem = 'source label appears in multiple plan rows';
    } else if (row.action !== 'Delete') {
      if (!row.name.trim() || !row.destination.trim()) {
        problem = 'proposed label is incomplete';
      } else if (row.destination !== reconstructed) {
        problem = 'proposed path does not match proposed parent and name';
      } else if (row.action === 'Keep' && row.current !== row.destination) {
        problem = 'path changed but action is Keep';
      } else if (row.action === 'Rename/Move' && row.current === row.destination) {
        problem = 'source and destination are the same';
      } else if (destinations.get(row.destination) > 1) {
        problem = 'duplicate proposed destination';
      } else if (row.action !== 'Keep' && livePaths.has(row.destination)) {
        problem = plannedSources.has(row.destination)
          ? 'destination already exists in Gmail'
          : 'stale plan: destination appeared in Gmail';
      }
    }

    if (!problem && /^(Adjusted:|Error:)/.test(row.validation)) {
      problem = 'review validation status and run Validate Label Plan again';
    }
    const requiredValidation = row.action === 'Delete' ? 'Review: deletion requested'
      : row.action === 'Keep' ? 'No change' : 'Ready';
    if (!problem && row.validation !== requiredValidation) {
      problem = 'run Validate Label Plan before preflight';
    }

    if (problem) {
      counts.blocked++;
      return ['Blocked: ' + problem];
    }
    if (row.action === 'Delete') {
      counts.deletion++;
      return ['Review: explicit deletion confirmation required'];
    }
    if (row.action === 'Keep') {
      counts.unchanged++;
      return ['No change'];
    }
    counts.ready++;
    return ['Ready: ' + row.action];
  });

  resultRange.setValues(results);
  ui.alert('Preflight complete. No Gmail changes were made.\n\n' +
    'Ready operations: ' + counts.ready + '\n' +
    'Deletion reviews: ' + counts.deletion + '\n' +
    'Blocked operations: ' + counts.blocked + '\n' +
    'Unchanged rows: ' + counts.unchanged + '\n\n' +
    (missing.length || added.length
      ? 'Stale Gmail state relative to plan: ' + missing.length +
        ' missing label(s), ' + added.length + ' added label(s).'
      : 'Live Gmail label paths match the plan snapshot.') + '\n' +
    'Preflight is a snapshot, not authorization to execute changes.');
}


function ensureRunLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Run Log';

  let log = ss.getSheetByName(sheetName);

  const headers = [
    'Timestamp',
    'User',
    'Action',
    'Original Label Path',
    'Requested Label Path',
    'Result',
    'Details',
    'Run ID'
  ];

  if (log) {
    // Append missing headers without moving columns or touching historical entries.
    const lastColumn = log.getLastColumn();
    const existing = lastColumn ? log.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
    const missing = headers.filter(header => !existing.includes(header));
    if (missing.length) {
      const requiredColumns = lastColumn + missing.length;
      if (requiredColumns > log.getMaxColumns()) {
        log.insertColumnsAfter(log.getMaxColumns(), requiredColumns - log.getMaxColumns());
      }
      log.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
    }
    return log;
  }

  log = ss.insertSheet(sheetName);

  log
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  log.setFrozenRows(1);

  log
    .getRange(1, 1, 1, headers.length)
    .setFontWeight('bold');

  return log;
}
