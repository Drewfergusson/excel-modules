const utils = require('./utils');

module.exports = {
  range,
  rangeFromCell
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

  let endingRow = rangeString.match(/:[A-Z]([0-9])+/)? Number(rangeString.match(/:[A-Z]+([0-9])+/)[1]): undefined;
  let endingColumn = rangeString.match(/:([A-Z])+/)? rangeString.match(/:([A-Z])+/)[1]: undefined;

  return { sheet, startingRow, startingColumn, endingRow, endingColumn };
}

 /**
  * Takes the start of the range and calculates a new range with a given height and width
  */
function rangeFromCell(rangeString, height, width) {
  const expandedRange = this.range(rangeString);
  expandedRange.addRowsToStartCell(height - 1).addColumnsToStartCell(width - 1);
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
    addRowsToStartCell,
    addColumnsToStartCell,
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

  function addRowsToStartCell(number) {
    this.endingRow = (this.startingRow) + number;
    return this;
  }

  function addColumnsToStartCell(number) {
    this.endingColumn = utils.columnAddition(this.startingColumn, number);
    return this;
  }
  /**
   * () => [Range]
   */
  function rows() {
    const rows = [];
    for(let row = startingRow; row <= endingRow; row++) {
      rows.push(range(`${this.sheet}!${this.startingColumn}${row}:${this.endingColumn}${row}`))
    }
    return rows;
  }
}
