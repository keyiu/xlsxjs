const defaultOption = {
  type: 'string',
};

module.exports = class Cell {
  /**
   * 构造方法
   * @method constructor
   * @param  {string/number}  value [description]
   * @param  {Object} option [description]
   */
  constructor(value, {type} = defaultOption) {
    this.value = value;
    this.type = type;
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
};
