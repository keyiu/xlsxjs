const assert = require('assert');
describe('workbook', () => {
  // 创建一个workbook
  describe('new', () => {
    it('创建一个workbook', (done) => {
      const Workbook = require('./workbook');
      let workbook = new Workbook();
      let sheet = workbook.addSheet('sheet1');
      const columns = [{
        title: '实收量',
        dataIndex: 'price',
        type: 'number',
        width: 13,
        render: (cell) => {
          if (cell.renderType === 2) {
            return cell.returnCount;
          }
          if (cell.renderType === 3) {
            return cell.realReplenishmentCount;
          }
          return cell.auditQuantity;
        },
      }];
      sheet.bulkRender({columns, dataSource: [{
        price: 1,
      }], hooks: {
        beforeRenderRow: (data) => {
        },
        beforeRenderCell: (data) => {
        },
      }});
      workbook.getBuffer().then((res) => {
        require('fs').writeFile('a.xlsx', res);
        done();
      }).catch((e) => {
        done(e);
      });
    }).timeout(5000);
  });
});
