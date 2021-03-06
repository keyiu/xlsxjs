const Cell = require('./cell');
module.exports = class Row {
  /**
   * 构造方法
   * @method constructor
   * @param  {Object}  worksheet [description]
   * @param  {string/number}  position [description]
   */
  constructor(worksheet, position) {
    this.position = position;
    this.worksheet = worksheet;
    this.pageBreak = false;
    this.cells = new Map();
  }
  /**
   * 添加一个分页字符
   * @method addPageBreak
   * @return {Object}
   */
  addPageBreak() {
    this.pageBreak = true;
    return this;
  }
  /**
   * 获取列中所有的单元格
   * @method getCells
   * @return {[type]} [description]
   */
  getCells() {
    return this.cells;
  }
  /**
   * 获取列中的单元格
   * @method getCell
   * @param  {[type]} position [description]
   * @return {[type]}          [description]
   */
  getCell(position) {
    return this.worksheet.getCell(Cell.getPosition(`${position}${this.position}`));
  }
  /**
   * [setCell description]
   * @method setCell
   * @param  {[type]} position [description]
   * @param  {[type]} cell     [description]
   */
  setCell(position, cell) {
    this.cells.set(position, cell);
  }
};
