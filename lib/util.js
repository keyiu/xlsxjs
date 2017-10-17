const xlsx = require('xlsx');
module.exports = function parseSheets(workbook) {
  workbook = _.reduce(workbook, (workbook, sheet, sheetName) => {
    sheet = _.reduce(sheet, (res, cell, key) => {
      // 去掉说明性字段
      if (key.match(/^!/)) {
        return res;
      }
      const {L, R} = splitLR(key);
      res[R] = res[R] || {};
      res[R][L] = getValue(cell);
      return res;
    }, {});
    workbook[sheetName] = sheet;
    return workbook;
  }, {});
  return workbook;
};

module.exports = function getValue({type, value}) {
  switch (type) {
    case 'd':
      return this.getDate(value);
    case 'b':
      return Boolean(value);
    case 'n':
      return Number(value);
    default:
      return value;
  }
};

module.exports = function getDate(value) {
  const {y: year, m: month, d: day, H: hour, M: minute, S: second} = xlsx.SSF.parse_date_code(value);
  return moment({year, month: month - 1, day, hour, minute, second});
};

module.exports = function getDate(value) {
  const {y: year, m: month, d: day, H: hour, M: minute, S: second} = xlsx.SSF.parse_date_code(value);
  return moment({year, month: month - 1, day, hour, minute, second});
};

/**
 * 拆分单元格所在的行列
 * @method splitColRow
 * @param  {[type]} str [description]
 * @return {[type]}     [description]
 */
module.exports.splitColRow = function splitColRow(str) {
  let col = str.match(/^[a-z|A-Z]+/gi);
	let row = str.match(/\d+$/gi);
  return {col, row};
}
