const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');

const code = readFileSync(join(__dirname, '..', 'Code.js'), 'utf8');

class Sheet {
  constructor(name, rows = []) {
    this.name = name;
    this.cells = new Map();
    this.formulas = new Map();
    this.validations = new Map();
    this.rules = [];
    this.maxRows = 20;
    this.writes = [];
    rows.forEach((row, r) => row.forEach((value, c) => {
      this.cells.set(`${r + 1},${c + 1}`, value);
    }));
  }
  getName() { return this.name; }
  clearContents() { this.cells.clear(); this.formulas.clear(); }
  getValue(r, c) {
    const formula = this.formulas.get(`${r},${c}`);
    if (formula === `=IF(E${r}="","",IF(D${r}="",E${r},D${r}&"/"&E${r}))`) {
      const parent = this.cells.get(`${r},4`);
      const name = this.cells.get(`${r},5`);
      return !name ? '' : !parent ? name : `${parent}/${name}`;
    }
    return this.cells.get(`${r},${c}`) ?? '';
  }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return 26; }
  getLastColumn() {
    return Math.max(0, ...[...this.cells, ...this.formulas]
      .filter(([, value]) => value !== '')
      .map(([key]) => Number(key.split(',')[1])));
  }
  getLastRow() {
    return Math.max(0, ...[...this.cells, ...this.formulas]
      .filter(([, value]) => value !== '')
      .map(([key]) => Number(key.split(',')[0])));
  }
  getRange(r, c, nr = 1, nc = 1) {
    assert.ok(r + nr - 1 <= this.maxRows);
    const sheet = this;
    const read = map => Array.from({ length: nr }, (_, i) =>
      Array.from({ length: nc }, (_, j) => map.get(`${r + i},${c + j}`) ?? ''));
    return {
      getSheet: () => sheet,
      getRow: () => r,
      getColumn: () => c,
      getLastRow: () => r + nr - 1,
      getLastColumn: () => c + nc - 1,
      getNumColumns: () => nc,
      getValues: () => Array.from({ length: nr }, (_, i) =>
        Array.from({ length: nc }, (_, j) => sheet.getValue(r + i, c + j))),
      getFormulas: () => read(sheet.formulas),
      clearContent() {
        sheet.writes.push(['clear', r, c, nr, nc]);
        for (let i = r; i < r + nr; i++) {
          for (let j = c; j < c + nc; j++) {
            sheet.cells.delete(`${i},${j}`);
            sheet.formulas.delete(`${i},${j}`);
          }
        }
        return this;
      },
      setValues(values) {
        sheet.writes.push(['values', r, c, nr, nc]);
        values.forEach((row, i) => row.forEach((value, j) => {
          // Sheets consumes one leading apostrophe as a text-entry escape.
          const stored = typeof value === 'string' && value.startsWith("'")
            ? value.slice(1) : value;
          sheet.cells.set(`${r + i},${c + j}`, stored);
        }));
        return this;
      },
      setFormula(formula) { sheet.formulas.set(`${r},${c}`, formula); return this; },
      clearDataValidations() { sheet.writes.push(['clearValidation', r, c, nr, nc]); },
      setDataValidation(rule) {
        sheet.writes.push(['validation', r, c, nr, nc]);
        for (let i = 0; i < nr; i++) sheet.validations.set(`${r + i},${c}`, rule);
      },
      setFontWeight() { sheet.writes.push(['font']); }
    };
  }
  insertRowsAfter(after, count) {
    assert.equal(after, this.maxRows);
    this.maxRows += count;
  }
  getConditionalFormatRules() { return this.rules; }
  setConditionalFormatRules(rules) { this.rules = rules; }
  setFrozenRows() { this.writes.push(['freeze']); }
}

function statusRule(sheet, formula = '=$H2="Ready"', color = '#e6f4ea', column = 1) {
  return {
    getBooleanCondition: () => ({
      getCriteriaType: () => 'CUSTOM_FORMULA',
      getCriteriaValues: () => [formula],
      getBackgroundObject: () => ({
        asRgbColor: () => ({ asHexString: () => color })
      })
    }),
    getRanges: () => [sheet.getRange(2, column, 3, 10)]
  };
}

function setup(sourceRows, plan, response = 'YES') {
  const sheets = new Map();
  if (sourceRows) sheets.set('Gmail Labels', new Sheet('Gmail Labels', sourceRows));
  if (plan) sheets.set('Label Plan', plan);
  const alerts = [];
  const ui = {
    Button: { YES: 'YES' }, ButtonSet: { YES_NO: 'YES_NO' },
    alert(...args) { alerts.push(args); return response; }
  };
  const context = vm.createContext({
    SpreadsheetApp: {
      flush() {},
      BooleanCriteria: { CUSTOM_FORMULA: 'CUSTOM_FORMULA' },
      getUi: () => ui,
      getActiveSpreadsheet: () => ({
        getSheetByName: name => sheets.get(name),
        insertSheet(name) { const sheet = new Sheet(name); sheets.set(name, sheet); return sheet; }
      }),
      newDataValidation: () => ({
        requireValueInList() { return this; },
        setAllowInvalid() { return this; }, build() { return this; }
      }),
      newConditionalFormatRule: () => ({
        whenFormulaSatisfied(formula) { this.formula = formula; return this; },
        setBackground(color) { this.color = color; return this; },
        setRanges(ranges) { this.ranges = ranges; return this; },
        build() {
          const rule = statusRule(this.ranges[0].getSheet(), this.formula, this.color);
          rule.getRanges = () => this.ranges;
          return rule;
        }
      })
    }
  });
  vm.runInContext(code, context);
  return { context, sheets, alerts };
}

const inventory = [['Full Label Path', 'Parent Path', 'Label Name'], ['Work', '', 'Work']];

test('ensureRunLogSheet_ creates headers once, returns the sheet, and never logs or alerts', () => {
  const { context, sheets, alerts } = setup(null);
  const log = context.ensureRunLogSheet_();
  assert.equal(log, sheets.get('Run Log'));
  assert.deepEqual(log.getRange(1, 1, 1, 8).getValues()[0],
    ['Timestamp', 'User', 'Action', 'Original Label Path', 'Requested Label Path', 'Result', 'Details', 'Run ID']);
  assert.equal(log.getLastRow(), 1);
  const writes = log.writes.length;
  assert.equal(context.ensureRunLogSheet_(), log);
  assert.equal(log.writes.length, writes);
  assert.equal(alerts.length, 0);
  assert.equal(context.createRunLog, undefined);
});

test('ensureRunLogSheet_ preserves historical rows and custom columns when adding Run ID', () => {
  const { context, sheets } = setup(null);
  const log = new Sheet('Run Log', [
    ['Timestamp', 'User', 'Action', 'Original Label Path', 'Requested Label Path', 'Result', 'Details', 'Custom'],
    ['Yesterday', 'User', 'Create', '', 'Work', 'Success', 'Details', 'Keep this']
  ]);
  sheets.set('Run Log', log);
  const original = [...log.cells];
  assert.equal(context.ensureRunLogSheet_(), log);
  original.forEach(([key, value]) => assert.equal(log.cells.get(key), value));
  assert.equal(log.getValue(1, 9), 'Run ID');
  assert.equal(log.getValue(2, 9), '');
  assert.deepEqual(log.writes, [['values', 1, 9, 1, 1]]);
  context.ensureRunLogSheet_();
  assert.equal(log.writes.length, 1);
});

test('menu no longer offers Run Log creation', () => {
  const { context } = setup(null);
  const items = [];
  const menu = {
    addItem(label, handler) { items.push([label, handler]); return this; },
    addSeparator() { return this; }, addToUi() {}
  };
  context.SpreadsheetApp.getUi().createMenu = () => menu;
  context.onOpen();
  assert.equal(items.some(([label, handler]) => /Run Log/.test(label) || /RunLog/.test(handler)), false);
  assert.ok(items.some(([, handler]) => handler === 'preflightLabelPlan'));
});

test('preparation repairs empty F and dropdowns idempotently without touching H:J or blank rows', () => {
  const plan = new Sheet('Label Plan', [['header'], ['', '', '', 'Parent', 'New', '', '', 'H', 'I', 'J'],
    [], ['', '', '', '', 'Other', 'Custom path', 'Delete']]);
  const { context } = setup(null, plan);
  context.prepareLabelPlanRows_();
  assert.equal(plan.getValue(2, 6), 'Parent/New');
  assert.equal(plan.getValue(2, 7), '');
  assert.deepEqual(plan.getRange(2, 8, 1, 3).getValues()[0], ['H', 'I', 'J']);
  assert.equal(plan.getValue(4, 6), 'Custom path');
  assert.equal(plan.getValue(4, 7), 'Delete');
  assert.ok(plan.validations.has('2,7'));
  assert.ok(plan.validations.has('4,7'));
  assert.equal(plan.validations.has('3,7'), false);
  const formulas = [...plan.formulas];
  const cells = [...plan.cells];
  context.prepareLabelPlanRows_();
  assert.deepEqual([...plan.formulas], formulas);
  assert.deepEqual([...plan.cells], cells);
});

test('validation adjusts pasted new/renamed/reverted rows once, preserves Delete and ignores blanks', () => {
  const plan = new Sheet('Label Plan', [['header'],
    ['', '', '', 'Parent', 'New'],
    ['Old', '', 'Old', '', 'Renamed', '', 'Keep'],
    ['Reverted', '', 'Reverted', '', 'Reverted', '', 'Rename/Move'],
    ['DeleteMe', '', 'DeleteMe', '', '', '', 'Delete'],
    [],
    ['Incomplete', '', 'Incomplete', '', '', '', 'Keep']]);
  const { context, alerts } = setup(null, plan);
  const userRule = statusRule(plan, '=TRUE');
  plan.rules = [userRule];
  context.validateLabelPlan();
  for (const [r, action] of [[2, 'Create'], [3, 'Rename/Move'], [4, 'Keep']]) {
    assert.equal(plan.getValue(r, 7), action);
    assert.match(plan.getValue(r, 8), /^Adjusted:/);
  }
  assert.equal(plan.getValue(5, 7), 'Delete');
  assert.equal(plan.getValue(5, 8), 'Review: deletion requested');
  assert.equal(plan.getValue(6, 8), '');
  assert.equal(plan.getValue(7, 7), 'Keep');
  assert.match(plan.getValue(7, 8), /^Error:/);
  assert.match(alerts.at(-1)[0], /Adjusted: 3/);
  assert.match(alerts.at(-1)[0], /Review adjusted rows/);
  assert.equal(plan.rules[0], userRule);
  const adjustedRule = plan.rules.find(rule =>
    rule.getBooleanCondition().getCriteriaValues()[0] === '=LEFT($H2,9)="Adjusted:"');
  assert.equal(adjustedRule.getBooleanCondition().getBackgroundObject().asRgbColor().asHexString(), '#fff4e5');
  context.validateLabelPlan();
  assert.equal(plan.getValue(2, 8), 'Ready');
  assert.equal(plan.getValue(3, 8), 'Ready');
  assert.equal(plan.getValue(4, 8), 'No change');
  assert.equal(plan.getValue(5, 8), 'Review: deletion requested');
  assert.match(alerts.at(-1)[0], /Adjusted: 0/);
  assert.equal(plan.rules.length, 6);
});

test('multi-row edit only clears stale statuses; validation later prepares proposals', () => {
  const plan = new Sheet('Label Plan', [['header'], ['', '', '', '', 'First'],
    ['', '', '', "'Parent", "'Second"]]);
  const { context } = setup(null, plan);
  plan.getRange(2, 8, 2, 3).setValues([['Ready', 'Old result', 'Old note'],
    ['Ready', 'Old result', 'Old note']]);
  plan.writes = [];
  plan.getLastRow = () => { throw new Error('onEdit must not scan the plan'); };
  context.onEdit({ range: plan.getRange(2, 4, 2, 2) });
  assert.deepEqual(plan.writes, [['clear', 2, 8, 2, 3]]);
  assert.equal(plan.formulas.size, 0);
  assert.equal(plan.validations.size, 0);
  assert.deepEqual(plan.getRange(2, 8, 2, 3).getValues(), [['', '', ''], ['', '', '']]);
  delete plan.getLastRow;
  context.validateLabelPlan();
  assert.equal(plan.getValue(2, 6), 'First');
  assert.equal(plan.getValue(3, 6), "'Parent/'Second");
  assert.equal(plan.getValue(2, 7), 'Create');
  assert.equal(plan.getValue(3, 7), 'Create');
  assert.ok(plan.validations.has('3,7'));
});

test('Delete preflight requires the exact deletion validation status', () => {
  const fixture = setupPreflight([preflightRow('Old', '', 'Delete')], ['Old']);
  for (const status of ['', 'Adjusted: test', 'Error: test', 'Ready', 'No change', 'Review: other']) {
    fixture.plan.cells.set('2,8', status);
    fixture.context.preflightLabelPlan();
    const expected = /^(Adjusted:|Error:)/.test(status)
      ? 'Blocked: review validation status and run Validate Label Plan again'
      : 'Blocked: run Validate Label Plan before preflight';
    assert.equal(fixture.plan.getValue(2, 9), expected);
    assert.equal(fixture.plan.getValue(2, 7), 'Delete');
    assert.equal(fixture.plan.getValue(2, 8), status);
  }
  fixture.context.validateLabelPlan();
  assert.equal(fixture.plan.getValue(2, 8), 'Review: deletion requested');
  fixture.context.preflightLabelPlan();
  assert.equal(fixture.plan.getValue(2, 9), 'Review: explicit deletion confirmation required');
});

test('preflight prepares but does not normalize; adjusted/error/unvalidated rows are blocked', () => {
  const fixture = setupPreflight([['', '', '', '', 'New']], []);
  fixture.context.preflightLabelPlan();
  assert.equal(fixture.plan.getValue(2, 6), 'New');
  assert.equal(fixture.plan.getValue(2, 7), '');
  assert.match(fixture.plan.getValue(2, 9), /^Blocked:/);
  fixture.context.validateLabelPlan();
  assert.match(fixture.plan.getValue(2, 8), /^Adjusted:/);
  fixture.context.preflightLabelPlan();
  assert.match(fixture.plan.getValue(2, 9), /^Blocked:/);
  fixture.context.validateLabelPlan();
  fixture.context.preflightLabelPlan();
  assert.equal(fixture.plan.getValue(2, 9), 'Ready: Create');
  for (const status of ['Error: test', 'Adjusted: test', '', 'Review: test']) {
    fixture.plan.cells.set('2,8', status);
    fixture.context.preflightLabelPlan();
    assert.match(fixture.plan.getValue(2, 9), /^Blocked:/);
    assert.equal(fixture.plan.getValue(2, 8), status);
  }
});

function preflightRow(current, destination, action) {
  const parts = destination.split('/');
  const name = parts.pop();
  return [current, '', '', parts.join('/'), name, destination, action,
    action === 'Keep' ? 'No change' : action === 'Delete' ? 'Review: deletion requested' : 'Ready',
    'Old preflight', 'User notes'];
}

function setupPreflight(rows, livePaths) {
  const plan = new Sheet('Label Plan', [['header'], ...rows]);
  const fixture = setup(null, plan);
  let calls = 0;
  // Only expose read APIs: any attempted Gmail mutation fails the test.
  fixture.context.GmailApp = {
    getUserLabels() {
      calls++;
      return livePaths.map(path => ({ getName: () => path }));
    }
  };
  return { ...fixture, plan, calls: () => calls };
}

for (const [title, current, destination, action, live, expected] of [
  ['source disappeared', 'Old', 'New', 'Rename/Move', [], 'Blocked: stale plan: source label no longer exists in Gmail'],
  ['Keep source disappeared', 'Old', 'Old', 'Keep', [], 'Blocked: stale plan: source label no longer exists in Gmail'],
  ['Delete source disappeared', 'Old', '', 'Delete', [], 'Blocked: stale plan: source label no longer exists in Gmail'],
  ['destination appeared', 'Old', 'New', 'Rename/Move', ['Old', 'New'], 'Blocked: stale plan: destination appeared in Gmail'],
  ['Create destination appeared', '', 'New', 'Create', ['New'], 'Blocked: stale plan: destination appeared in Gmail'],
  ['valid Create', '', 'New', 'Create', [], 'Ready: Create'],
  ['valid Rename/Move', 'Old', 'New', 'Rename/Move', ['Old'], 'Ready: Rename/Move'],
  ['Delete review', 'Old', '', 'Delete', ['Old'], 'Review: explicit deletion confirmation required'],
  ['Keep unchanged', 'Old', 'Old', 'Keep', ['Old'], 'No change'],
  ['Keep edited', 'Old', 'New', 'Keep', ['Old'], 'Blocked: path changed but action is Keep'],
  ['Rename unchanged', 'Old', 'Old', 'Rename/Move', ['Old'], 'Blocked: source and destination are the same'],
  ['apostrophe unchanged', "'Overdue", "'Overdue", 'Keep', ["'Overdue"], 'No change'],
  ['nested apostrophe rename', "xFlags/'Overdue", "'Waiting", 'Rename/Move', ["xFlags/'Overdue"], 'Ready: Rename/Move'],
  ['invalid action', 'Old', 'Old', 'Invalid', ['Old'], 'Blocked: invalid or blank action'],
  ['Create with source', 'Old', 'New', 'Create', ['Old'], 'Blocked: Create requires a blank current label path'],
  ['blank source', '', 'New', 'Rename/Move', [], 'Blocked: source label is blank'],
  ['blank destination', '', '', 'Create', [], 'Blocked: proposed label is incomplete']
]) {
  test(`preflight: ${title}`, () => {
    const { context, plan, calls } = setupPreflight([preflightRow(current, destination, action)], live);
    context.preflightLabelPlan();
    assert.equal(calls(), 1);
    assert.equal(plan.getValue(2, 9), expected);
    assert.equal(plan.getValue(2, 8), preflightRow(current, destination, action)[7]);
    assert.equal(plan.getValue(2, 10), 'User notes');
    assert.ok(plan.writes.every(([kind, , column, , width]) =>
      (column === 9 || (kind === 'validation' && column === 7)) && width === 1));
  });
}

test('preflight blocks duplicate destinations and occupied destinations scheduled for deletion', () => {
  const duplicate = setupPreflight([
    preflightRow('Old', 'New', 'Rename/Move'), preflightRow('', 'New', 'Create')
  ], ['Old']);
  duplicate.context.preflightLabelPlan();
  assert.equal(duplicate.plan.getValue(2, 9), 'Blocked: duplicate proposed destination');
  assert.equal(duplicate.plan.getValue(3, 9), 'Blocked: duplicate proposed destination');
  const occupied = setupPreflight([
    preflightRow('Old', 'New', 'Rename/Move'), preflightRow('New', '', 'Delete')
  ], ['Old', 'New']);
  occupied.context.preflightLabelPlan();
  assert.equal(occupied.plan.getValue(2, 9), 'Blocked: destination already exists in Gmail');
});

test('preflight detects tampered paths and duplicate sources independently of H', () => {
  const row = preflightRow('', 'New', 'Create');
  row[5] = 'Other';
  const fixture = setupPreflight([row, preflightRow('Old', 'A', 'Rename/Move'),
    preflightRow('Old', '', 'Delete')], ['Old']);
  fixture.context.preflightLabelPlan();
  assert.match(fixture.plan.getValue(2, 9), /does not match/);
  assert.match(fixture.plan.getValue(3, 9), /multiple plan rows/);
  assert.match(fixture.plan.getValue(4, 9), /multiple plan rows/);
});

test('preflight refreshes live state each run and summarizes stale state and all outcomes', () => {
  const live = ['Keep', 'Rename', 'Delete', 'Disappearing'];
  const fixture = setupPreflight([
    preflightRow('Keep', 'Keep', 'Keep'),
    preflightRow('Rename', 'Renamed', 'Rename/Move'),
    preflightRow('', 'Created', 'Create'),
    preflightRow('Delete', '', 'Delete'),
    preflightRow('Disappearing', 'Disappearing', 'Keep')
  ], live);
  fixture.context.preflightLabelPlan();
  assert.equal(fixture.plan.getValue(6, 9), 'No change');
  live.splice(live.indexOf('Disappearing'), 1);
  live.push('Unrelated new label');
  fixture.context.preflightLabelPlan();
  assert.equal(fixture.calls(), 2);
  assert.match(fixture.plan.getValue(6, 9), /source label no longer exists/);
  const summary = fixture.alerts.at(-1)[0];
  assert.match(summary, /Ready operations: 2\nDeletion reviews: 1\nBlocked operations: 1\nUnchanged rows: 1/);
  assert.match(summary, /Stale Gmail state relative to plan: 1 missing label\(s\), 1 added label\(s\)/);
});

test('preflight clears old results on empty rows and leaves unrelated cells and rules alone', () => {
  const fixture = setupPreflight([preflightRow('Keep', 'Keep', 'Keep')], ['Keep']);
  fixture.plan.cells.set('4,9', 'Old ready');
  fixture.plan.cells.set('5,10', 'Notes only');
  fixture.plan.cells.set('8,11', 'Outside');
  const rules = [statusRule(fixture.plan)];
  fixture.plan.rules = rules;
  fixture.context.preflightLabelPlan();
  assert.equal(fixture.plan.getValue(4, 9), '');
  assert.equal(fixture.plan.getValue(5, 10), 'Notes only');
  assert.equal(fixture.plan.getValue(8, 11), 'Outside');
  assert.equal(fixture.plan.rules, rules);
  assert.match(fixture.alerts.at(-1)[0], /Blocked operations: 0\nUnchanged rows: 1/);
});

test('preflight requires a plan and clears stale approval if the live read fails', () => {
  const absent = setup(null);
  absent.context.preflightLabelPlan();
  assert.equal(absent.alerts[0][0], 'Please create the Label Plan first.');
  const empty = setup(null, new Sheet('Label Plan', [['header']]));
  empty.context.preflightLabelPlan();
  assert.equal(empty.alerts[0][0], 'There are no rows to preflight.');
  const fixture = setupPreflight([preflightRow('', 'New', 'Create')], []);
  fixture.context.GmailApp.getUserLabels = () => { throw new Error('Unavailable'); };
  fixture.context.preflightLabelPlan();
  assert.equal(fixture.plan.getValue(2, 9), 'Blocked: live Gmail labels could not be read');
  assert.equal(fixture.plan.getValue(2, 8), 'Ready');
  assert.equal(fixture.plan.getValue(2, 10), 'User notes');
});

for (const labelPath of ["'Overdue", "'To post", "'Waiting", "xFlags/'Overdue",
  "'Parent/Child", "'Parent/'Overdue", "''Overdue", 'Work', 'Work/Projects', "O'Brien"]) {
  test(`export, regenerate, and validate preserve literal label text: ${labelPath}`, () => {
    const { context, sheets } = setup(null);
    context.GmailApp = {
      getUserLabels: () => [{
        getName: () => labelPath,
        getThreads: () => [1, 2],
        getUnreadCount: () => 1
      }]
    };
    const parts = labelPath.split('/');
    const name = parts.pop();
    const parent = parts.join('/');
    for (let run = 0; run < 2; run++) {
      context.exportGmailLabels();
      const source = sheets.get('Gmail Labels');
      assert.deepEqual(source.getRange(2, 1, 1, 6).getValues()[0],
        [labelPath, parent, name, labelPath.split('/').length, 2, 1]);
      context.createLabelPlan();
      const plan = sheets.get('Label Plan');
      assert.deepEqual(plan.getRange(2, 1, 1, 7).getValues()[0],
        [labelPath, parent, name, parent, name, labelPath, 'Keep']);
      context.validateLabelPlan();
      assert.equal(plan.getValue(2, 8), 'No change');
    }
  });
}

test('missing, header-only, and blank-path inventories leave the plan untouched', () => {
  for (const source of [null, [inventory[0]], [inventory[0], ['', '', 'stray']]]) {
    const plan = new Sheet('Label Plan', [['header'], ['proposal']]);
    const { context, sheets } = setup(source, plan);
    context.createLabelPlan();
    assert.equal(plan.writes.length, 0);
    assert.equal(plan.cells.get('2,1'), 'proposal');
    const fresh = setup(source);
    fresh.context.createLabelPlan();
    assert.equal(fresh.sheets.has('Label Plan'), false);
    assert.equal(sheets.get('Label Plan'), plan);
  }
});

test('No and closing the confirmation preserve values and formula-only planning work', () => {
  for (const response of ['NO', 'CLOSE']) {
    for (const formulaOnly of [false, true]) {
      const plan = new Sheet('Label Plan');
      if (formulaOnly) plan.formulas.set('2,6', '=IF(TRUE,"","")');
      else plan.cells.set('2,10', 'Keep my notes');
      const { context, alerts } = setup(inventory, plan, response);
      context.createLabelPlan();
      assert.equal(alerts[0][0], 'Regenerate Label Plan?');
      assert.equal(plan.writes.length, 0);
    }
  }
});

test('regeneration preserves outside content and unrelated rules, and is repeatable', () => {
  const plan = new Sheet('Label Plan', [['header'], ['old'], ['obsolete']]);
  plan.cells.set('15,11', 'outside');
  plan.formulas.set('16,12', '=1+1');
  const userRules = [
    statusRule(plan, '=TRUE'),
    statusRule(plan, '=$H2="Ready"', '#ffffff'),
    statusRule(plan, '=$H2="Ready"', '#e6f4ea', 11),
    { getBooleanCondition: () => null }
  ];
  plan.rules = [statusRule(plan), ...userRules];
  const { context } = setup(inventory, plan);
  for (let run = 0; run < 2; run++) {
    context.createLabelPlan();
    assert.equal(plan.rules.length, 9);
    userRules.forEach((rule, i) => assert.equal(plan.rules[i], rule));
    assert.equal(plan.cells.get('15,11'), 'outside');
    assert.equal(plan.formulas.get('16,12'), '=1+1');
    assert.equal(plan.cells.get('2,7'), 'Keep');
    assert.equal(plan.cells.has('3,1'), false);
    assert.equal(plan.formulas.get('2,6'), '=IF(E2="","",IF(D2="",E2,D2&"/"&E2))');
    assert.ok(plan.writes.every(([, , c, , nc]) => c >= 1 && c + nc - 1 <= 10));
    assert.ok(plan.writes.filter(([kind]) => kind.includes('Validation') || kind === 'validation')
      .every(([, , c, , nc]) => c === 7 && nc === 1));
  }
});

test('outside-only content does not prompt; new plans are initialized and can grow', () => {
  const plan = new Sheet('Label Plan');
  plan.cells.set('2,11', 'outside');
  const existing = setup(inventory, plan);
  existing.context.createLabelPlan();
  assert.equal(existing.alerts.length, 1);
  const fresh = setup([inventory[0], ...Array.from({ length: 25 }, (_, i) => [`L${i}`, '', `L${i}`])]);
  fresh.sheets.get('Gmail Labels').maxRows = 26;
  fresh.context.createLabelPlan();
  const created = fresh.sheets.get('Label Plan');
  assert.equal(created.maxRows, 26);
  assert.ok(created.writes.some(([kind]) => kind === 'font'));
  assert.ok(created.writes.some(([kind]) => kind === 'freeze'));
});

test('onEdit clears H:J only for data rows intersecting D:G, including multi-cell pastes', () => {
  const { context } = setup(inventory);
  context.onEdit();
  for (const [name, r, c, nr, nc, expected] of [
    ['Label Plan', 2, 4, 1, 1, [[2, 8, 1, 3]]],
    ['Label Plan', 3, 7, 3, 1, [[3, 8, 3, 3]]],
    ['Label Plan', 1, 2, 4, 8, [[2, 8, 3, 3]]],
    ['Label Plan', 2, 6, 1, 1, [[2, 8, 1, 3]]],
    ['Label Plan', 1, 4, 1, 4, []],
    ['Label Plan', 2, 1, 3, 3, []],
    ['Label Plan', 2, 8, 3, 4, []],
    ['Gmail Labels', 2, 4, 3, 4, []]
  ]) {
    const sheet = new Sheet(name);
    context.onEdit({ range: sheet.getRange(r, c, nr, nc) });
    assert.deepEqual(sheet.writes.map(write => write.slice(1)), expected);
  }
});
