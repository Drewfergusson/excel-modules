'use strict';

var test = require('tape');
var range = require('./range');

test('testing sheet of created range', function (t) {
  t.plan(1);
  var cells = range('Sheet1!A5');
  t.equal(cells.sheet, 'Sheet1');
});

test('creating a range from values', function (t) {
  t.plan(6);
  var values = [[1, 2, 4], ['A', 'B', 2]];
  var cells = range().from('Sheet1!A5').values(values);
  t.equal(cells.sheet, 'Sheet1');

  var span = range().from('Sheet1!A1').dimensions({ rows: 3, columns: 10 });
  t.equal(span.location(), 'Sheet1!A1:J3');

  t.comment('Testing start and end of range values');
  t.equal(cells.start().toString(), 'A5');
  t.equal(cells.end().toString(), 'C6');

  t.comment('Testing range width and height');
  t.equal(cells.height(), 2);
  t.equal(cells.width(), 3);
});