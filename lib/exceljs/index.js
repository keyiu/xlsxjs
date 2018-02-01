const Excel = require('exceljs');
const Cell = require('../cell');
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
  if (data.pageSetup) {
    console.log(data.pageSetup);
    sheet.pageSetup = data.pageSetup;
  }
  data.rows.forEach((row) => {
    return _renderRow(sheet, row);
  });
  data.cells.forEach((cell, key) => {
    const [startPosition, endPosition] = key.split(':');
    _renderCell(sheet, startPosition, cell);
    if (endPosition) {
      sheet.mergeCells(key);
    }
  });
}
/**
 * 渲染工作表
 * @method      _render
 * @param       {[type]} sheet [description]
 * @param       {[type]} row  [description]
 * @constructor
 */
function _renderRow(sheet, {position, pageBreak, height}) {
  let row = sheet.getRow(position);
  if (pageBreak) {
    row.addPageBreak();
  }
  if (height) {
    row.height = height;
  }
}
/**
 * 渲染工作表
 * @method      _renderCell
 * @param       {[type]} sheet [description]
 * @param       {[type]} position  [description]
 * @param       {[type]} options  [description]
 * @constructor
 */
function _renderCell(sheet, position, {value, style = {}}) {
  let cell = sheet.getCell(position);
  cell.value = value;
  if (style.font) {
    cell.font = style.font;
  }
  if (style.alignment) {
    cell.alignment = style.alignment;
  }
  if (style.border) {
    cell.border = style.border;
  }
  if (style.numberFormat) {
    cell.numFmt = style.numberFormat;
  }
  if (style.fill) {
    cell.fill = style.fill;
  }
  if (style.width) {
    let {startCol, endCol} = Cell.getRowCol(position);
    let startColNum = Cell.parse10(startCol);
    let endColNum = Cell.parse10(endCol);
    if ((endColNum - startColNum) === 0) {
      sheet.getColumn(startCol).width = style.width;
    } else {
      let width = parseInt(style.width / (endColNum - startColNum));
      for (let i = 0; i < endColNum - startColNum; i++) {
        sheet.getColumn(Cell.parse26(startColNum + 1 + i)).width = width;
      }
    }
    sheet.getColumn(startCol);
  }
}
