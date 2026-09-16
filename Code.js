function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gmail Label Manager')
    .addItem('Export Gmail Labels', 'exportGmailLabels')
    .addItem('Create Label Plan', 'createLabelPlan')
    .addItem('Validate Label Plan', 'validateLabelPlan')
    .addSeparator()
    .addItem('Create Run Log', 'createRunLog')
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
      .setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    'Export complete. ' +
    rows.length +
    ' user-created Gmail labels were found.'
  );
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
    .setValues(rows);

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
  const actionRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(
      [
        'Keep',
        'Rename/Move',
        'Create',
        'Delete'
      ],
      true
    )
    .setAllowInvalid(false)
    .build();

  plan
    .getRange(2, 7, rows.length, 1)
    .setDataValidation(actionRule);

  /*
   * Conditional formatting managed by the script.
   * Ordinary user formatting and column widths remain untouched.
   */
  const dataRange = plan.getRange(
    2,
    1,
    rows.length,
    headers.length
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
      .build()
  ];

  const userRules = plan.getConditionalFormatRules()
    .filter(rule => !isLabelPlanStatusRule_(rule));
  plan.setConditionalFormatRules(userRules.concat(rules));

  if (isNewPlan) {
    plan.setFrozenRows(1);
    plan.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  SpreadsheetApp.getUi().alert(
    'Label Plan created with ' +
    rows.length +
    ' existing Gmail labels.\n\n' +
    'No Gmail changes were made.'
  );
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
    '=LEFT($H2,6)="Review"': '#fff4e5'
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

  const rowCount = lastRow - 1;

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

  const results = data.map(row => {
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
    'Errors: ' + errors
  );
}


function createRunLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Run Log';

  let log = ss.getSheetByName(sheetName);

  if (log) {
    SpreadsheetApp.getUi().alert(
      'Run Log already exists.'
    );
    return;
  }

  log = ss.insertSheet(sheetName);

  const headers = [
    'Timestamp',
    'User',
    'Action',
    'Original Label Path',
    'Requested Label Path',
    'Result',
    'Details'
  ];

  log
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  log.setFrozenRows(1);

  log
    .getRange(1, 1, 1, headers.length)
    .setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    'Run Log created.\n\n' +
    'No Gmail changes were made.'
  );
}
