const Cell = require('./cell');
const Row = require('./row');
const _ = require('lodash');

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
    this.maxUsedCol = 0;
    this.maxUsedRow = 0;
    this.pageSetup = {};
    this.cells = new Map();
    this.rows = new Map();
    this.cols = new Map();
    this.name = name;
  }
  /**
   * 获取一行
   * @method getRow
   * @param  {number} position [description]
   * @return {[type]}          [description]
   */
  getRow(position) {
    if (typeof this.rows.get(position) !== 'object') {
      this.rows.set(position, new Row(position));
    }
    return this.rows.get(position);
  }
  /**
   * 获取单元格（如果没有则创建）
   * @method getCell
   * @param  {String} positionStr 单元格位置
   * @return {Object}          单元格对象
   */
  getCell(positionStr) {
    const position = Cell.getPosition(positionStr);
    if (typeof this.cells.get(position) !== 'object') {
      this.lastUsedCell = new Cell(position);
      let {endCol, endRow} = Cell.getRowCol(position);
      this.lastUsedCol = endCol;
      this.lastUsedRow = endRow;
      this.cells.set(position, this.lastUsedCell);
    }
    return this.cells.get(position);
  }
  /**
   * 设置value
   * @method setValue
   * @param  {[type]} v [description]
   * @return {[type]}   [description]
   */
  setValue(v) {
    this.value = v;
    return this;
  }
  /**
   * 设置单元格的值
   * @method setCell
   * @param  {[type]} positionStr [description]
   * @param  {[type]} value    [description]
   * @return {Object}           单元格对象
   */
  setCell(positionStr, value) {
    let cell = this.getCell(positionStr);
    cell.value = value;
    return cell;
  }
  /**
   * 批量渲染工作表中的数据
   * @method bulkRender
   * @return {Object}
   */
  bulkRender({dataSource, columns, hooks, titleStyle}) {
    this.renderTitle(columns, titleStyle);
    this.renderBody(columns, dataSource);
    return this;
  }
  /**
   * 渲染表头
   * @method renderTitle
   * @param  {[type]}    columns   [description]
   * @param  {[type]}    style   [description]
   */
  renderTitle(columns, style) {
    // 为列添加一些默认处理
    _.forEach(columns, (column, key) => {
      column.index = column.index || Cell.parse26(key + 1);
      column.render = column.render || _defaultRender;
    });
    let jumpTitle = _.isEmpty(_.filter(columns, (column) => {
      return column.title;
    }));
    if (jumpTitle) {
      return;
    }
    let rowNumber = this.lastUsedRow + 1;
    for (let i = 0; i < columns.length; i++) {
      let column = columns[i];
      let index = column.index;
      let position = Cell.getPosition(`${index}${rowNumber}`);
      this.getCell(position).setValue(column.title).setStyle(_.assign({width: column.width}, style));
    }
  }
  /**
   * 渲染表体
   * @method renderBody
   * @param  {[type]}   columns    [description]
   * @param  {[type]}   dataSource [description]
   */
  renderBody(columns, dataSource) {
    const that = this;
    _.forEach(dataSource, (data) => {
      let row = that.lastUsedRow + 1;
      _.forEach(columns, (column) => {
        let cell = that.getCell(Cell.getPosition(`${row}${column.index}`));
        let value = column.render(data);
        cell.setValue(value);
        if (column.style) {
          cell.setStyle(column.style);
        }
      });
    });
  }
};
/**
 * 默认的渲染方法
 * @method defaultRender
 * @param  {[type]}      data [description]
 * @return {[type]}           [description]
 */
function _defaultRender(data) {
  return data[this.dataIndex];
}
