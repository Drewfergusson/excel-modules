const test = require('tape');
const range = require('./range');

test('testing sheet of created range', t => {
  t.plan(1);
  const cells = range('Sheet1!A5');
  t.equal(cells.sheet, 'Sheet1');
});

test('creating a range from values', t => {
  t.plan(3);
  const values = [
    [1, 2, 4], ['A', 'B', 2]
  ];
  const cells = range().from('Sheet1!A5').values(values);
  console.log(cells.values());
  t.equal(cells.sheet, 'Sheet1');
  t.equal(cells.start().toString(), 'A5');
  t.equal(cells.end().toString(), 'C6')
});