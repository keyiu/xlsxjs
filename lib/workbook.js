const xlsx = require('xlsx');
const Worksheet = require('./worksheet');
const {renderBuffer} = require('./exceljs');
const Cell = require('./cell');
const _ = require('lodash');
module.exports = class Workbook {
  /**
   * 构造方法
   * @method constructor
   */
  constructor() {
    this.sheets = new Map();
  }
  /**
   * 获取工作薄中的工作表
   * @method getSheet
   * @param  {String} name 表名称
   * @return {Object}      工作表
   */
  getWorksheet(name) {
    if (!this.sheets.get(name)) {
      this.addSheet(name);
    }
    return this.sheets.get(name);
  }
  /**
   * 创建一个工作表
   * @method addSheet
   * @param  {[type]} name [description]
   * @return {Object}     添加一个工作表
   */
  addWorksheet(name) {
    const worksheet = new Worksheet('name');
    this.sheets.set(name, worksheet);
    return worksheet;
  }
  /**
   * 获取工作薄对象的buffer
   * @method getBuffer
   * @return {[type]}  [description]
   */
  writeToBuffer() {
    let buffer = renderBuffer(this.sheets);
    return buffer;
  }

  /**
   * 解析工作表
   * @method parse
   * @param  {[type]} buffer [description]
   * @return {[type]}        [description]
   */
  static parse(buffer) {
    const workbook = new this();
    let sheets = xlsx.read(buffer).Sheets;
    _.forEach(sheets, (sheet, sheetName) => {
      let ws = workbook.addWorksheet(sheetName);
      _.forEach(sheet, (cell, position) => {
        if (Cell.isPosition(position)) {
          ws.getCell(position).setValue(cell.v);
        }
      });
    });
    return workbook;
  }
};
