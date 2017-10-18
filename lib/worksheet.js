const Cell = require('./cell');
const {getPosition} = require('./util');
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
    this.cells = new Map();
    this.rows = new Map();
    this.cols = new Map();
    this.name = name;
  }
  /**
   * 获取单元格（如果没有则创建）
   * @method getCell
   * @param  {String} positionStr 单元格位置
   * @return {Object}          单元格对象
   */
  getCell(positionStr) {
    const position = getPosition(positionStr);
    if (typeof this.cells.get(position) !== 'object') {
      this.cells.set(position, new Cell());
    }
    return this.cells.get(position);
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
   */
  bulkRender({dataSource, columns, hooks}) {
    this.renderTitle(columns);
    this.renderBody(columns, dataSource);
  }
  /**
   * 渲染表头
   * @method renderTitle
   * @param  {[type]}    columns   [description]
   */
  renderTitle(columns) {
    // 为列添加一些默认处理
    _.forEach(columns, (column, key) => {
      column.index = column.index || Cell.parse26(key + 1);
      column.render = _defaultRender;
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
      let position = getPosition(`${index}${rowNumber}`);
      this.getCell(position).value(column.title);
    }
  }
  /**
   * 渲染表体
   * @method renderBody
   * @param  {[type]}   columns    [description]
   * @param  {[type]}   dataSource [description]
   */
  renderBody(columns, dataSource) {
    let rowNumber = this.lastUsedRow + 1;
    const that = this;
    _.forEach(dataSource, (data) => {
      _.forEach(columns, (column) => {
        let cell = that.getCell(getPosition(`${that.lastUsedRow}${column.index}`));
        let value = column.render(data);
        cell.value(value);
      });
    });
    for (let i = 0; i < columns.length; i++) {
      let column = columns[i];
      let index = column.index || Cell.parse26(i + 1);
      let position = getPosition(`${index}${rowNumber}`);
      this.getCell(position).value(column.title);
    }
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
