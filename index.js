const cell = require('./modules/cell');
const range = require('./modules/range');

module.exports = {
  range,
  cell
}

 /**
  * Takes the start of the range and calculates a new range with a given height and width
  */
function rangeFromCell(rangeString, height, width) {
  const expandedRange = this.range(rangeString);
  expandedRange.addRowsToStartCell(height - 1).addColumnsToStartCell(width - 1);
  return expandedRange;
}


