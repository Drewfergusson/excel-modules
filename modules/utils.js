'use strict';
// setting undefined because the JS compliler evaluates (null || '') as null
// also using .split('') because ...String also needs a polyfill on IE which (I think) is the version
// of iframe browser used on Windows Excel
const lettersArr = [undefined, ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')];
const lettersInAlphabet = 26;/**

/**
 *
 * The non-zeroth based column from the given starting column
 * @return {String}
 */
function nthColumnFrom(startingColumn, n) {
  return getColumnFromIndex(getIndexFromColumn(startingColumn) + (n - 1));
}

/**
 * The non-zeroth based row form the given starting row
 * @return
 */
function nthRowFrom(startingRow, n) {
  return startingRow + (n - 1);
}

/**
 * Using the starging column, get the ending column for a row of a given length
 */
function rowEndingColumn(startingColumn, rowLength) {
  return getColumnFromIndex(getIndexFromColumn(startingColumn) + rowLength);
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

function sheet(rangeString) {
  if(!rangeString) {
    return 'Sheet1'
  }
  const sheet = rangeString.match(/(.+)!/)? rangeString.match(/(.+)!/)[1]: undefined;
  if(!sheet) {
    throw new Error('Unable to parse sheet');
  }
  return sheet;
}

function range(rangeString) {
  if(!rangeString) {
    return { startingColumn: 'A', startingRow: 1, endingRow: 1, endingColumn: 'A' }
  }
  const startingRow = rangeString.match(/!*[A-Z]+([0-9]+):*/)? Number(rangeString.match(/!*[A-Z]+([0-9]+):*/)[1]): undefined;
  const startingColumn =  rangeString.match(/!*([A-Z]+)[0-9]+:*/)? rangeString.match(/!*([A-Z]+)[0-9]+:*/)[1]: undefined;
  if(!startingRow || !startingColumn) {
    throw new Error('Unable to parse start of range');
  }
  const endingRow = rangeString.match(/:[A-Z]([0-9]+)/)? Number(rangeString.match(/:[A-Z]+([0-9]+)/)[1]): startingRow;
  const endingColumn = rangeString.match(/:([A-Z]+)/)? rangeString.match(/:([A-Z]+)/)[1]: startingColumn;
  return { startingColumn, startingRow, endingRow, endingColumn };
}

module.exports = {
  rowEndingColumn,
  getIndexFromColumn,
  getColumnFromIndex,
  nthColumnFrom,
  nthRowFrom,
  parse: { sheet, range }
}
