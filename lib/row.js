const _ = require('lodash');
const defaultOption = {
  type: 'string',
};

module.exports = class Row {
  /**
   * 构造方法
   * @method constructor
   * @param  {string/number}  position [description]
   */
  constructor(position) {
    this.position = position;
    this.pageBreak = false;
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
};
