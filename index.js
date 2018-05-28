// setting undefined because the JS compliler evaluates (null || '') as null
// also using .split('') because ...String also needs a polyfill on IE which (I think) is the version
// of iframe browser used on Windows Excel
const lettersArr = [undefined, ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')];
const lettersInAlphabet = 26;

module.exports = {
  range: rangeFactory,
  cell: cellFactory
}

/**
 *
 * ({startingColumn: String, addedColumns: Number}) => String
 * ('A', 5) => 'F'
 * ('AA' + 1) => 'AB'
 */
function columnAddition(startingColumn, addedColumns) {
  return getColumnFromIndex(getIndexFromColumn(startingColumn) + addedColumns);
}

/**
 * Using the starging column, get the ending column for a row of a given length
 */
function rowEndingColumn(startingColumn, rowLength) {
  // addedColumns -1 to get row that includes the starting column in the total length
  return getColumnFromIndex(getIndexFromColumn(startingColumn) + rowLength - 1);
}

/**
 * (column: String) => Number
 * ('F') => 5
 * ('AA') => 27
 */
function getIndexFromColumn(column) {
  return column.split('').reduce((acc, currentVal) => {
    return lettersInAlphabet * acc + lettersArr.indexOf(currentVal);
  }, 0);
}

// TODO, this needs tests and will likely fail for a third letter
/**
 * (index: Number) => String
 * (5) => 'F'
 * (27) => 'AA'
 */
function getColumnFromIndex(index) {
  return (lettersArr[Math.floor(index / lettersInAlphabet)] || '') + lettersArr[(index % lettersInAlphabet)];
}

/**
 * @typedef {Object} Range
 * @param {function} toString
 * @param {String} startingColumn
 * @param {String} endingColumn
 * @param {Number} startingRow
 * @param {Number} endingRow
 */

/**
 * @example
 * const selectedRange = range('A1:I1')
 * const fullRange = range('')
 */
 function rangeFactory(rangeString) {
  let sheet = rangeString.match(/(.+)!/)?  rangeString.match(/(.+)!/)[1] : undefined;

  let startingRow = rangeString.match(/[A-Z]+([0-9])+:/)? Number(rangeString.match(/[A-Z]+([0-9])+:/)[1]): undefined;
  let startingColumn =  rangeString.match(/([A-Z]+)[0-9]+:/)? rangeString.match(/([A-Z]+)[0-9]+:/)[1] : undefined;

  let endingRow = endOfRange.match(/:[A-Z]([0-9])+/)? Number(endOfRange.match(/:[A-Z]([0-9])+/)[1]): undefined;
  let endingColumn = endOfRange.match(/:([A-Z])+/)? endOfRange.match(/:([A-Z])+/)[1]: undefined;

  return {
    toString: () => {
      return `${startingColumn}${startingRow}:${endingColumn || ''}:${endingRow || ''}`
    },
    startOfRange: `${startingColumn}${startingRow}`,
    endOfRange: `${endingColumn}:${endingRow}`,
    startingRow,
    endingRow,
    startingColumn,
    endingColumn,
    sheet,
    getLocation: () => {
      return `${sheet || ''}!${startingColumn}${startingRow}:${endingColumn || ''}:${endingRow || ''}`
    },
    addRowsDown: (number) => {
      endingRow = endingRow + number;
      return this;
    },
    addColumnsRight: (number) => {
      endingColumn = columnAddition(endingColumn, number);
      return this;
    }
  };
}

/**
 *
 */
function cellFactory(cellString) {
  let sheet = rangeString.match(/(.+)!/)?  rangeString.match(/(.+)!/)[1] : '';

  let row = Number(cellString.match(/[0-9]+/)[0]);
  let column = startOfRange.match(/[A-Z]+/)[0];

  return {
    row,
    column,
    toString: `${column}${row}`,
    sheet,
    location: `${sheet}!${column}${row}`,
    addRowsDown: (number) => {
      return rangeFactory(`${sheet}!${column}${row}:${column}${row + number}`);
    },
    addColumnsRight: (number) => {
      return rangeFactory(`${sheet}!${column}${row}:${columnAddition(column, number)}${row}`);
    }
  };
}
