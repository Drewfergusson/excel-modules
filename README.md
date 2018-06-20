# Accell

A library created to more easily work with the Excel Javascript API; mainly through the range object
which is a query-able and easily manipulable object that can be created from a range string ||
a starting cell and a 2D array of values.

## Installation

```bash
npm install https://github.com/Drewfergusson/excel-modules.git
```

## Usage
```js
  const range = require('excel-modules');

  const selectedRange = range('A1:H5');
  selectedRange.start().row; // => 1
  selectedRange.end().column; // => 'H'

  const otherRange = range('Sheet1!A1:H5');
  otherRange.sheet; // => 'Sheet1'
  otherRange.toString(); // => 'A1:H5'
  otherRange.location() ;// => 'Sheet1!A1:H5'

  const rangeFromValues = range('Sheet2!A1').from.values([
    ['a', 'b', 'c'], [1, 2, 3]
  ]);
  rangeFromValue.toString(); // => 'Sheet2:A1:B3'
```
