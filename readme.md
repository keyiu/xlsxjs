# 一个简单的excel管理库

## 安装

```shell
npm i xlsxjs --save
```

## 使用方法

```javascript
const Xlsx = require('xlsxjs');
const workbook = new Xlsx();
let worksheet = workbook.addWorksheet('sheetName');
// 样式写法可以参考 https://github.com/guyonroche/exceljs#styles
const titleStyle = {
    font: {
      bold: true,
      size: 9,
      name: 'SimSun',
    },
}
// 渲染单个单元格
worksheet
    .getCell(`A1`)
    .setValue('a1')
    .setStyle(titleStyle);

// 合并单元格
worksheet
    .getCell(`A2:B2`)
    .setValue('a2:b2')
    .setStyle(titleStyle);
// 批量渲染单元格
worksheet.bulkRender({
    dataSource: [{
        title: 'title',
        name: '实例'
    }],
    worksheet,
    columns: [{
      title: '商品编号',
      dataIndex: 'skuCode',
      style: bodyStyle,
      width: 10,
    },
    {
      title: '商品名称',
      style: bodyStyle,
      dataIndex: 'name',
      width: 31,
    }],
    titleStyle,
  });
//  渲染buffer
let buffer = yield workbook.writeToBuffer();
return buffer;
```

## 说明
该库是对exceljs，xlsx的二次封装，其渲染时使用的exceljs作为民渲染引擎，解析时是使用xlsx作为解析引擎使用的。这里只是提供了一个更加容易书写的方式而已。