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