'use strict';

var test = require('tape');
var utils = require('./utils');

test('test gettingIndexFromColumn', function (t) {
  t.plan(3);
  t.equal(utils.getIndexFromColumn('A'), 1);
  t.equal(utils.getIndexFromColumn('AA'), 27);
  t.equal(utils.getIndexFromColumn('AB'), 28);
});

test('test nthColumnFrom', function (t) {
  t.plan(3);
  t.equal(utils.nthColumnFrom('A', 1), 'A');
  t.equal(utils.nthColumnFrom('A', 4), 'D');
  t.equal(utils.nthColumnFrom('B', 4), 'E');
});

test('Parsing a range string', function (t) {
  t.plan(5);
  var rangeString = 'Sheet1!A7:AA9';
  var sheet = utils.parse.sheet(rangeString);

  var _utils$parse$range = utils.parse.range(rangeString),
      startingColumn = _utils$parse$range.startingColumn,
      startingRow = _utils$parse$range.startingRow,
      endingColumn = _utils$parse$range.endingColumn,
      endingRow = _utils$parse$range.endingRow;

  t.equal(sheet, 'Sheet1');
  t.equal(startingColumn, 'A');
  t.equal(startingRow, 7);
  t.equal(endingColumn, 'AA');
  t.equal(endingRow, 9);
});