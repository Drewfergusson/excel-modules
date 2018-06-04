const test = require('tape');
const range = require('./range');

test('testing sheet of created range', (t) => {
  t.plan(1);
  const cells = range('Sheet1!A5');
  t.equal(cells.sheet, 'Sheet1');
});