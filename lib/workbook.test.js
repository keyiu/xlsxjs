const assert = require('assert');
describe('workbook', () => {
  // 创建一个workbook
  describe('new', () => {
    it('创建一个workbook', (done) => {
      const Workbook = require('./workbook');
      let workbook = new Workbook();
      let sheet = workbook.addSheet('sheet1');
      workbook.getBuffer().then((res) => {
        console.log(res, '89890890890890890890');
        require('fs').writeFile('a.xlsx', res);
        done();
      }).catch((e) => {
        done(e);
      });
    });
  });
});
