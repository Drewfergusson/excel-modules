const test = require('tape');
const utils = require('./utils');

test('test gettingIndexFromColumn', t => {
  t.plan(3);
  t.equal(utils.getIndexFromColumn('A'), 1);
  t.equal(utils.getIndexFromColumn('AA'), 27);
  t.equal(utils.getIndexFromColumn('AB'), 28);
});

test('test nthColumnFrom', t => {
  t.plan(3);
  t.equal(utils.nthColumnFrom('A', 1), 'A');
  t.equal(utils.nthColumnFrom('A', 4), 'D');
  t.equal(utils.nthColumnFrom('B', 4), 'E');
});

test('Parsing a range string', t => {
  t.plan(5);
  const rangeString = 'Sheet1!A7:AA9';
  const sheet = utils.parse.sheet(rangeString)
  const { startingColumn, startingRow, endingColumn, endingRow } = utils.parse.range(rangeString);
  t.equal(sheet, 'Sheet1')
  t.equal(startingColumn, 'A');
  t.equal(startingRow, 7);
  t.equal(endingColumn, 'AA');
  t.equal(endingRow, 9);
});