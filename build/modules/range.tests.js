'use strict';

var test = require('tape');
var range = require('./range');

test('testing sheet of created range', function (t) {
  t.plan(1);
  var cells = range('Sheet1!A5');
  t.equal(cells.sheet, 'Sheet1');
});

test('creating a range from values', function (t) {
  t.plan(5);
  var values = [[1, 2, 4], ['A', 'B', 2]];
  var cells = range().from('Sheet1!A5').values(values);
  t.equal(cells.sheet, 'Sheet1');

  t.comment('Testing start and end of range values');
  t.equal(cells.start().toString(), 'A5');
  t.equal(cells.end().toString(), 'C6');

  t.comment('Testing range width and height');
  t.equal(cells.height(), 2);
  t.equal(cells.width(), 3);
});