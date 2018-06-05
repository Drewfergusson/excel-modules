'use strict';
const utils = require('./utils');

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
 * const selectedRange = range('Sheet1!A1:I1')
 * const fullRange = range('Master!A5')
 */
const range = rangeString => {
  let sheet = utils.parse.sheet(rangeString);
  let { startingRow, startingColumn, endingRow, endingColumn } = utils.parse.range(rangeString);
  let rangeValues = [];

  /**
   * @return {String}
   */
  const rangeStartString = () => `${startingColumn}:${startingRow}`;

  /**
   * @return {String}
   */
  const rangeEndString = () => `${endingColumn}:${endingRow}`;

  /**
   * @return {String}
   */
  const toString = () => {
    if(endingColumn && endingRow) {
      return `${startingColumn}${startingRow}:${endingColumn || ''}${endingRow || ''}`;
    }
    return `${startingColumn}${startingRow}`;
  }

  /**
   * @return {String}
   */
  const getLocation = () => {
    if(!endingColumn || !endingColumn) {
      return `${sheet + '!' || ''}${startingColumn}${startingRow}`;
    }
    return `${sheet + '!' || ''}${startingColumn}${startingRow}:${endingColumn}${endingRow}`;
  }

  /**
   * @return {Range}
   */
  const addRowsDown = number => {
    endingRow = (endingRow) + number;
    return;
  }

  /**
   * @return {Range}
   */
  const addColumnsRight = number => {
    endingColumn = utils.columnAddition(endingColumn, number);
    return;
  }

  /**
   * @return {Number}
   */
  const height = () => {
    return endingRow - startingRow;
  }

  /**
   * @return {Number}
   */
  const width = () => {
    return utils.getIndexFromColumn(endingColumn) - utils.getIndexFromColumn(startingColumn);
  }

  /**
   * @return {Range}
   */
  const startCell = () => {
    return range(`${sheet}!${startingColumn}${startingRow}`);
  }
  /**
   * @return {Boolean}
   */
  const isCell = () => {
    return startingRow === endingRow && startingColumn === endingColumn
  }
  /**
   * Rerturns the rows of the total range as an array of individual rows
   * () => [Range]
   */
  const rows = () => {
    const rows = [];
    for(let row = startingRow; row <= endingRow; row++) {
      rows.push(range(`${sheet}!${startingColumn}${row}:${endingColumn}${row}`));
    }
    return rows;
  }

  const addValues = values => {
    if(values.length !== height()) {
      throw new Error(`${values.length} rows passed in won't fit cleanly into this range of ${height()} rows`);
    }
    values.forEach((row, index) => {
      if(row.length !== width()) {
        throw new Error(`Row ${index} with a width of ${row.length} will not fit cleanly into this range of width ${width()}`)
      }
    })
    return rangeValues = values;
  }

  const values = () => {
    return rangeValues;
  }

  const from = rangeString => {
    let newRange = range(rangeString).startCell();
    return {
      values: values => {
        values.reduce((acc, row) => {
          if(row.length !== acc) {
            throw new Error('Rows are not all of the same length');
          }
          return row.length
        }, values[0].length);
        newRange.addColumnsRight(values[0].length);
        newRange.addRowsDown(values.length);
        newRange.addValues(values);
        return newRange;
      }
    };
  }

  return {
    start: () => ({ row: startingRow, column: startingColumn, toString: rangeStartString }),
    end: () => ({ row: endingRow, column: endingColumn, toString: rangeEndString }),
    toString, sheet, getLocation, addRowsDown, addColumnsRight, rows, height, startCell,
    addValues, width, values, from
  };
}

module.exports = range;

