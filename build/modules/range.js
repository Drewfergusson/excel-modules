'use strict';

var utils = require('./utils');

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
var range = function range(rangeString) {
  var sheet = utils.parse.sheet(rangeString);

  var _utils$parse$range = utils.parse.range(rangeString),
      startingRow = _utils$parse$range.startingRow,
      startingColumn = _utils$parse$range.startingColumn,
      endingRow = _utils$parse$range.endingRow,
      endingColumn = _utils$parse$range.endingColumn;

  var rangeValues = [];

  /**
   * @return {String}
   */
  var rangeStartString = function rangeStartString() {
    return '' + startingColumn + startingRow;
  };

  /**
   * @return {String}
   */
  var rangeEndString = function rangeEndString() {
    return '' + endingColumn + endingRow;
  };

  /**
   * @return {String}
   */
  var toString = function toString() {
    if (endingColumn && endingRow) {
      return '' + startingColumn + startingRow + ':' + (endingColumn || '') + (endingRow || '');
    }
    return '' + startingColumn + startingRow;
  };

  /**
   * @return {String}
   */
  var location = function location() {
    if (!endingColumn || !endingColumn) {
      return '' + (sheet + '!' || '') + startingColumn + startingRow;
    }
    return '' + (sheet + '!' || '') + startingColumn + startingRow + ':' + endingColumn + endingRow;
  };

  var setWidth = function setWidth(width) {
    endingColumn = utils.nthColumnFrom(endingColumn, width);
    return;
  };

  var setHeight = function setHeight(height) {
    endingRow = utils.nthRowFrom(endingRow, height);
    return;
  };

  /**
   * @return {Number}
   */
  var height = function height() {
    return endingRow - (startingRow - 1);
  };

  /**
   * @return {Number}
   */
  var width = function width() {
    return utils.getIndexFromColumn(endingColumn) - (utils.getIndexFromColumn(startingColumn) - 1);
  };

  /**
   * @return {Range}
   */
  var startCell = function startCell() {
    return range(sheet + '!' + startingColumn + startingRow);
  };
  /**
   * @return {Boolean}
   */
  var isCell = function isCell() {
    return startingRow === endingRow && startingColumn === endingColumn;
  };
  /**
   * Rerturns the rows of the total range as an array of individual rows
   * () => [Range]
   */
  var rows = function rows() {
    var rows = [];
    for (var row = startingRow; row <= endingRow; row++) {
      rows.push(range(sheet + '!' + startingColumn + row + ':' + endingColumn + row));
    }
    return rows;
  };

  var addValues = function addValues(values) {
    if (values.length !== height()) {
      throw new Error(values.length + ' rows passed in won\'t fit cleanly into this range of ' + height() + ' rows');
    }
    values.forEach(function (row, index) {
      if (row.length !== width()) {
        throw new Error('Row ' + index + ' with a width of ' + row.length + ' will not fit cleanly into this range of width ' + width());
      }
    });
    return rangeValues = values;
  };

  var values = function values() {
    return rangeValues;
  };

  var from = function from(rangeString) {
    var newRange = range(rangeString).startCell();
    return {
      values: function values(_values) {
        _values.reduce(function (acc, row) {
          if (row.length !== acc) {
            throw new Error('Rows are not all of the same length');
          }
          return row.length;
        }, _values[0].length);
        newRange.setWidth(_values[0].length);
        newRange.setHeight(_values.length);
        newRange.addValues(_values);
        return newRange;
      },
      dimensions: function dimensions(_ref) {
        var rows = _ref.rows,
            columns = _ref.columns;

        newRange.setHeight(rows);
        newRange.setWidth(columns);
        return newRange;
      }
    };
  };

  return {
    start: function start() {
      return { row: startingRow, column: startingColumn, toString: rangeStartString };
    },
    end: function end() {
      return { row: endingRow, column: endingColumn, toString: rangeEndString };
    },
    toString: toString, sheet: sheet, location: location, rows: rows, height: height, startCell: startCell, addValues: addValues, width: width, values: values, from: from, setWidth: setWidth,
    setHeight: setHeight
  };
};

module.exports = range;