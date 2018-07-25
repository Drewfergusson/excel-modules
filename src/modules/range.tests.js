const test = require('tape');
const range = require('./range');

test('testing sheet of created range', t => {
  t.plan(1);
  const cells = range('Sheet1!A5');
  t.equal(cells.sheet, 'Sheet1');
});

test('creating a range from values', t => {
  t.plan(6);
  const values = [
    [1, 2, 4], ['A', 'B', 2]
  ];
  const cells = range().from('Sheet1!A5').values(values);
  t.equal(cells.sheet, 'Sheet1');

  const span = range().from('Sheet1!A1').dimensions({rows: 3, columns: 10});
  t.equal(span.location(), 'Sheet1!A1:J3');

  t.comment('Testing start and end of range values');
  t.equal(cells.start().toString(), 'A5');
  t.equal(cells.end().toString(), 'C6');

  t.comment('Testing range width and height');
  t.equal(cells.height(), 2);
  t.equal(cells.width(), 3);
});

test('sheet name edge cases', t => {
  t.plan(2);

  const spaceInName = range("'Sheet 1'!:A5");
  const tabInName = range("'Sheet\t1'!:A5");

  t.equal(spaceInName.sheet, "Sheet 1");
  t.equal(tabInName.sheet, "Sheet\t1")

});