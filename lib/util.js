const xlsx = require('xlsx');
const _ = require('lodash');
const moment = require('moment');
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
  const {
    y: year,
    m: month,
    d: day,
    H: hour,
    M: minute,
    S: second,
  } = xlsx.SSF.parse_date_code(value);
  return moment({year, month: month - 1, day, hour, minute, second});
};

/**
 * 拆分单元格所在的行列
 * @method splitColRow
 * @param  {[type]} str [description]
 * @param  {[type]} row [description]
 * @return {[type]}     [description]
 */
module.exports.getPosition = function getPosition(str, row) {
  let [startStr, endStr] = str.split(':');
  let positionArr = [];
  let startCol = startStr.match(/^[a-z|A-Z]+/gi);
  let startRow = row || startStr.match(/\d+$/gi);
  if (startCol && startRow) {
    positionArr.push(`${startCol}${startRow}`);
  }
  if (endStr) {
    let endCol = endStr.match(/^[a-z|A-Z]+/gi);
    let endRow = row || endStr.match(/\d+$/gi);
    if (endCol && endRow) {
      positionArr.push(`${endCol}${endRow}`);
    }
  }
  return positionArr.join(':');
};
