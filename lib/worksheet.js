const exceljs = require('exceljs')
const Cell = require('./cell')
const {splitColRow} = require('./util');

module.exports = class Worksheet {
  /**
   * 构造方法
   * @method constructor
   * @param  {String}    name 表名称
   */
  constructor(name) {
    this.lastUsedRow = 0;
    this.lastUsedCol = 0;
    this.lastUsedCell = 0;
    this.cells = new Map();
    this.rows = new Map();
    this.cols = new Map();
    this.name = name;
    this.mergedCells = new Set();
  }
  /**
   * 获取单元格（如果没有则创建）
   * @method getCell
   * @param  {String} position 单元格位置
   * @return {Object}          单元格对象
   */
  getCell(position) {
    const {col, row} = splitColRow(position);
    if (!(this.cells.get(col) instanceof Map)) {
      this.cells.set(col, new Map());
    }
    if (!(this.cells.get(col).get(row) instanceof Map)) {
      this.cells.get(col).set(row, {});
    }
    return this.cells.get(col).get(row);
  }
  /**
   * 设置单元格的值
   * @method setCell
   * @param  {[type]} position [description]
   * @param  {[type]} value    [description]
   * @return {Object}           单元格对象
   */
  setCell(position, value) {
    const {col, row} = splitColRow(position);
    if (!(this.cells.get(col) instanceof Map)) {
      this.cells.set(col, new Map());
    }
    if (!(this.cells.get(col).get(row) instanceof Map)) {
      this.cells.get(col).set(row, {});
    }
    this.cells.get(col).get(row).value = value;
    return this.cells.get(col).get(row);
  }
};
