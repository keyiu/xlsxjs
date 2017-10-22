const Cell = require('./cell');
module.exports = class Col {
  /**
   * 构造方法
   * @method constructor
   * @param  {Object}  worksheet [description]
   * @param  {string/number}  position [description]
   */
  constructor(worksheet, position) {
    this.position = position;
    this.worksheet = worksheet;
    this.cells = new Map();
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
