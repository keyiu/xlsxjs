const _ = require('lodash');
const defaultOption = {
  type: 'string',
};

module.exports = class Cell {
  /**
   * 构造方法
   * @method constructor
   * @param  {string/number}  position [description]
   */
  constructor(position) {
    this.position = position;
  }
  /**
   * [parse26 description]
   * @method parse26
   * @param  {[type]} n [description]
   * @return {[type]}         [description]
   */
  static parse26(n) {
    let s = '';
    while (n > 0) {
      let m = n % 26;
      if (m === 0) m = 26;
      s = String.fromCharCode(m + 64) + s;
      n = (n - m) / 26;
    }
    return s;
  }
  /**
   * 设置单元格的值
   * @method setValue
   * @param  {[type]} value [description]
   * @return {Object}
   */
  setValue(value) {
    this.value = value;
    return this;
  }
  /**
   * 设置单元格的样式
   * @method setStyle
   * @param  {[type]} style [description]
   * @return {Object}
   */
  setStyle(style) {
    this.style = _.assign(this.style, style);
    return this;
  }
  /**
   * 获取单元格的坐标
   * @method getPosition
   * @param  {[type]}    str [description]
   * @param  {[type]}    row [description]
   * @return {[type]}        [description]
   */
  static getPosition(str, row) {
    let [startStr, endStr] = str.split(':');
    let positionArr = [];
    let [startCol] = startStr.match(/[A-Z]+/gi);
    let [startRow] = row || startStr.match(/\d+/gi);
    if (startCol && startRow && _.toNumber(startRow) !== 0) {
      positionArr.push(`${startCol}${startRow}`);
    }
    if (endStr) {
      let [endCol] = endStr.match(/[A-Z]+/gi);
      let [endRow] = row || endStr.match(/\d+/gi);
      if (endCol && endRow && _.toNumber(endRow) !== 0) {
        positionArr.push(`${endCol}${endRow}`);
      }
    }
    if (_.isEmpty(positionArr)) {
      throw new Error(`${str}${row}${startRow}${startCol}`);
    }
    return positionArr.join(':');
  }
  /**
   * 获取单元格所在的行列
   * @method getRowCol
   * @param  {[type]}    position [description]
   * @return {[type]}        [description]
   */
  static getRowCol(position) {
    let [startStr, endStr] = position.split(':');
    let [startCol] = startStr.match(/[A-Z]+/gi);
    let [startRow] = startStr.match(/\d+/gi);
    let res = {startRow: _.toNumber(startRow), endRow: _.toNumber(startRow), startCol, endCol: startCol};
    if (endStr) {
      let [endCol] = endStr.match(/[A-Z]+/gi);
      let [endRow] = endStr.match(/\d+/gi);
      res.endCol = endCol;
      res.endRow = _.toNumber(endRow);
    }
    return res;
  }
};
