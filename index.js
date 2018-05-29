// setting undefined because the JS compliler evaluates (null || '') as null
// also using .split('') because ...String also needs a polyfill on IE which (I think) is the version
// of iframe browser used on Windows Excel
const lettersArr = [undefined, ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')];
const lettersInAlphabet = 26;

module.exports = {
  range,
  rangeFromCell
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
 * Returns an object with the raw values of a parsed range string
 */
function parseRange(rangeString) {
  let sheet = rangeString.match(/(.+)!/)?  rangeString.match(/(.+)!/)[1] : undefined;

  let startingRow = rangeString.match(/!*[A-Z]+([0-9])+:*/)? Number(rangeString.match(/!*[A-Z]+([0-9])+:*/)[1]): undefined;
  let startingColumn =  rangeString.match(/!*([A-Z]+)[0-9]+:*/)? rangeString.match(/!*([A-Z]+)[0-9]+:*/)[1] : undefined;

  if(!startingRow && !startingColumn) {
    throw new Error('Invalid range string');
  }

  let endingRow = rangeString.match(/:[A-Z]([0-9])+/)? Number(endOfRange.match(/:[A-Z]([0-9])+/)[1]): undefined;
  let endingColumn = rangeString.match(/:([A-Z])+/)? endOfRange.match(/:([A-Z])+/)[1]: undefined;

  return { sheet, startingRow, startingColumn, endingRow, endingColumn };
}

 /**
  * Takes the start of the range and calculates a new range with a given height and width
  */
function rangeFromCell(rangeString, height, width) {
  const expandedRange = this.range(rangeString);
  expandedRange.addRowsDown(height - 1).addColumnsRight(width - 1);
  return expandedRange;
}

/**
 * @example
 * const selectedRange = range('A1:I1')
 * const fullRange = range('')
 */
 function range(rangeString) {
  let { sheet, startingRow, startingColumn, endingRow, endingColumn } = parseRange(rangeString);

  return {
    toString,
    startOfRange,
    endOfRange,
    startingRow,
    endingRow,
    startingColumn,
    endingColumn,
    sheet,
    getLocation,
    addRowsDown,
    addColumnsRight,
    rows
  };

  function startOfRange() {
    return `${this.startingColumn}${this.startingRow}`;
  }

  function endOfRange() {
    return `${this.endingColumn}${this.endingRow}`;
  }

  function toString() {
    if(this.endingColumn && this.endingRow) {
      return `${this.startingColumn}${this.startingRow}:${this.endingColumn || ''}${this.endingRow || ''}`;
    }
    return `${this.startingColumn}${this.startingRow}`;
  }

  function getLocation() {
    return `${this.sheet || ''}!${this.startingColumn}${this.startingRow}:${this.endingColumn || ''}:${this.endingRow || ''}`
  }

  function addRowsDown(number) {
    this.endingRow = (this.endingRow || this.startingRow)+ number;
    return this;
  }

  function addColumnsRight(number) {
    this.endingColumn = columnAddition(this.endingColumn || this.startingColumn, number);
    return this;
  }
  /**
   * () => [Range]
   */
  function rows() {
    const rows = [];
    for(let row = startingRow; row > endingRow; row++) {
      rows.push[range(`${this.sheet}!${this.startingColumn}${row}:${this.endingColumn}${row}`)]
    }
    return rows;
  }
}
