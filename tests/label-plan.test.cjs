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
    this.rules = [];
    this.maxRows = 20;
    this.writes = [];
    rows.forEach((row, r) => row.forEach((value, c) => {
      this.cells.set(`${r + 1},${c + 1}`, value);
    }));
  }
  getName() { return this.name; }
  getMaxRows() { return this.maxRows; }
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
      getValues: () => read(sheet.cells),
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
          sheet.cells.set(`${r + i},${c + j}`, value);
        }));
        return this;
      },
      setFormula(formula) { sheet.formulas.set(`${r},${c}`, formula); return this; },
      clearDataValidations() { sheet.writes.push(['clearValidation', r, c, nr, nc]); },
      setDataValidation() { sheet.writes.push(['validation', r, c, nr, nc]); },
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
    assert.equal(plan.rules.length, 8);
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
