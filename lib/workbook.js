const Excel = require('exceljs');
const xlsx = require('xlsx');
const moment = require('moment');
const Worksheet = require('./worksheet');
const {getDate, parseSheets} = require('./util');
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
  getSheet(name) {
    if (!this.sheets.get(name)) {
      this.addSheet(name);
    }
    return this.sheets.get(name);
  }
  /**
   * 创建一个工作表
   * @method addSheet
   * @param  {[type]} name [description]
   */
  addSheet(name) {
    const worksheet = new Worksheet('name');
    this.sheets.set(name, worksheet);
  }
  /**
   * 获取工作薄对象的buffer
   * @method getBuffer
   * @return {[type]}  [description]
   */
  getBuffer() {
    const workbook = new Excel.Workbook();
    let sheet = workbook.addWorksheet('asdf');
    let row = sheet.getRow(1);
    row.getCell(1).value = 5;
    // this.sheets.forEach((sheet, key) => {
    //   console.log(key, '9880089089');
    //   // let sheet = workbook.addWorksheet(key);
    // });
    return workbook.xlsx.writeBuffer();
  }
  /**
   * 解析工作表
   * @method parse
   * @param  {[type]} buffer [description]
   * @return {[type]}        [description]
   */
  static parse(buffer) {
    let workbook = xlsx.read(buffer).Sheets;
    workbook = parseSheets(workbook);
    return workbook;
  }
};
