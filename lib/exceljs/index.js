const Excel = require('exceljs');

module.exports.renderBuffer = function(sheets) {
  const workbook = _renderWorkbook();
  sheets.forEach((data, key) => {
    let sheet = workbook.addWorksheet(key);
    _renderWorksheet(sheet, data);
  });
  return workbook.xlsx.writeBuffer();
};

/**
 * 创建工作表对象
 * @method      _renderWorkbook
 * @constructor
 * @return      {[type]}        [description]
 */
function _renderWorkbook() {
  const workbook = new Excel.Workbook();
  return workbook;
}
/**
 * 渲染工作表
 * @method      _render
 * @param       {[type]} sheet [description]
 * @param       {[type]} data  [description]
 * @constructor
 */
function _renderWorksheet(sheet, data) {
  data.cells.forEach((cell, key) => {
    const [startPosition, endPosition] = key.split(':');
    sheet.getCell(startPosition).value = cell.value;
    if (endPosition) {
      sheet.mergeCells(key);
    }
  });
}
