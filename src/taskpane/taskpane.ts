/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("addCluster").onclick = addCluster;
    document.getElementById("addRegister").onclick = addRegister;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      // const range = context.workbook.getSelectedRange();

      // // Read the range address
      // range.load("address");

      // // Update the fill color
      // range.format.fill.color = "green";

      //range.rowIndex
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
      ]);

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      await context.sync();
      // console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function addRegister() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Create the headers and format them to stand out.
      let headers = [
        ["register"]
      ];
      range.load("rowIndex");
      const usedRange = sheet.getUsedRange();
      usedRange.load('columnCount');
      await context.sync();

      // usedRange.columnCount;
      // range.rowIndex;
      const row_start = range.rowIndex + 1;
      let row = row_start;

      const columnCount = 14 > usedRange.columnCount ? 14 : usedRange.columnCount;
      //const columnCount = 14;

      const columnstr = String.fromCharCode(65 + columnCount - 1);
      let rangestr = "A" + row + ":" + columnstr + row;
      let headerRange = sheet.getRange(rangestr);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);

      rangestr = "A" + row_start + ":N" + (row_start + 8);
      const bodyRange = sheet.getRange(rangestr);
      bodyRange.conditionalFormats.clearAll();
      bodyRange.dataValidation.clear();
      bodyRange.format.borders.getItem('InsideHorizontal').style = 'Continuous';
      bodyRange.format.borders.getItem('InsideVertical').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeRight').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeTop').style = 'Continuous';

      bodyRange.format.borders.getItem('EdgeTop').weight = 'Medium';
      bodyRange.format.borders.getItem('EdgeBottom').weight = 'Medium';
      bodyRange.format.borders.getItem('EdgeLeft').weight = 'Medium';
      bodyRange.format.borders.getItem('EdgeRight').weight = 'Medium';
      

      rangestr = 'A' + row
      headerRange = sheet.getRange(rangestr);
      headerRange.values = headers;

      rangestr = "A" + row + ":N" + row;
      headerRange = sheet.getRange(rangestr);
      headerRange.merge();

      headerRange.format.fill.color = "#D0CECE";
      headerRange.format.font.color = "Black";
      headerRange.format.font.bold = true;
      headerRange.format.horizontalAlignment = 'Left';

      headers = [["name", "addressOffset", "size", "access", "resetValue", "resetMask",
        "", "headRegisterName", "alternateRegister", "alternateGroupName", "dim", "dimIncrement", "dimName", "description"]
      ];
      row += 1;
      rangestr = "A" + row + ":N" + row;
      headerRange = sheet.getRange(rangestr);
      headerRange.values = headers;
      headerRange.format.fill.color = "#D0CECE";
      headerRange.format.font.color = "Black";
      headerRange.format.font.bold = true;

      headers = [["fields:", "", "", "", "", "range", "", "enumValues", "", "", "", "", ""],
      ["bitRange", "bitName", "access", "defaultValue", "writeConstraint", "minimum", "maximum", "enumName", "enumValue", "enumDescription", "description", "", "pathHDL"]
      ];
      row += 2
      rangestr = "B" + row + ":N" + (row + 1);
      headerRange = sheet.getRange(rangestr);
      headerRange.values = headers;
      headerRange.format.fill.color = "#D0CECE";
      headerRange.format.font.color = "Black";
      headerRange.format.font.bold = true;




      rangestr = "A" + (row_start + 3) + ":A" + (row_start + 8);
      let merge_range = sheet.getRange(rangestr);
      merge_range.merge();
      rangestr = "B" + (row_start + 3) + ":F" + (row_start + 3);
      merge_range = sheet.getRange(rangestr);
      merge_range.merge();
      rangestr = "G" + (row_start + 3) + ":H" + (row_start + 3);
      merge_range = sheet.getRange(rangestr);
      merge_range.merge();
      rangestr = "I" + (row_start + 3) + ":K" + (row_start + 3);
      merge_range = sheet.getRange(rangestr);
      merge_range.merge();


      rangestr = "A" + (row_start + 2);
      let dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "register的标识字符串",
        showPrompt: true,
        title: "name"
      };
      let textLenRule = {
        textLength: {
          formula1: 16,
          operator: Excel.DataValidationOperator.lessThan
        }
      };
      dataRange.dataValidation.rule = textLenRule;

      rangestr = "B" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "相对于器件的“baseAddress”的register地址偏移。是0x开始的字串",
        showPrompt: true,
        title: "addressOffset"
      };
      let textLenBetwRule = {
        textLength: {
          formula1: 4,
          formula2: 6,
          operator: Excel.DataValidationOperator.between
        }
      };
      dataRange.dataValidation.rule = textLenBetwRule;

      rangestr = "C" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "器件中包含的寄存器的默认位宽",
        showPrompt: true,
        title: "size"
      };
      let listRule = {
        list: {
          source: "8,16,32,64",
          inCellDropDown: true
        }
      };
      dataRange.dataValidation.rule = listRule;

      rangestr = "D" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "定义寄存器的默认访问权限",
        showPrompt: true,
        title: "access"
      };
      listRule = {
        list: {
          source: "R,W,RW",
          inCellDropDown: true
        }
      };
      dataRange.dataValidation.rule = listRule;

      rangestr = "E" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "定义复位后所有寄存器的默认值",
        showPrompt: true,
        title: "resetValue"
      };
      textLenRule = {
        textLength: {
          formula1: 10,
          operator: Excel.DataValidationOperator.equalTo
        }
      };
      dataRange.dataValidation.rule = textLenRule;

      rangestr = "F" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "标识那些定义了复位值的寄存器位",
        showPrompt: true,
        title: "resetMask"
      };
      dataRange.dataValidation.rule = textLenRule;

      rangestr = "I" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "引用上面已定义的寄存器到当前位置的描述中",
        showPrompt: true,
        title: "alternate"
      };

      rangestr = "K" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "定义寄存器数组，并指定元素数量",
        showPrompt: true,
        title: "registerDim"
      };
      let wholeRule = {
        wholeNumber: {
          formula1: 2,
          formula2: 32,
          operator: Excel.DataValidationOperator.between
        }
      };
      dataRange.dataValidation.rule = wholeRule;


      rangestr = "L" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "指定两个相邻元素之间的地址增量（以字节为单位）",
        showPrompt: true,
        title: "dimIncrement"
      };

      rangestr = "M" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "指定C类型结构的名称。如果未定义，则使用寄存器name",
        showPrompt: true,
        title: "dimName"
      };

      rangestr = "N" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      dataRange.dataValidation.prompt = {
        message: "描述寄存器详细信息的字符串",
        showPrompt: true,
        title: "description"
      };

      rangestr = "A" + (row_start + 2) + ':F' + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      let conditionalFormat = dataRange.conditionalFormats.add(
        Excel.ConditionalFormatType.custom
      );

      // Set the font of negative numbers to red.
      conditionalFormat.custom.format.borders.bottom.color = "red";
      conditionalFormat.custom.format.borders.top.color = "red";
      conditionalFormat.custom.format.borders.left.color = "red";
      conditionalFormat.custom.format.borders.right.color = "red";
      conditionalFormat.custom.format.fill.color = '#FFC3C3';

      conditionalFormat.priority = 0;
      conditionalFormat.stopIfTrue = true;

      conditionalFormat.custom.rule.formula = '=ISBLANK(A' + (row_start + 2) + ')';

      rangestr = "B" + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      conditionalFormat = dataRange.conditionalFormats.add(
        Excel.ConditionalFormatType.custom
      );

      // Set the font of negative numbers to red.
      conditionalFormat.custom.format.font.color = 'red';
      conditionalFormat.custom.format.fill.color = '#F4B084';
      conditionalFormat.custom.rule.formula = '= NOT(LEFT(B' + (row_start + 2) + ',2) ="0x")';

      rangestr = "E" + (row_start + 2) + ':F' + (row_start + 2);
      dataRange = sheet.getRange(rangestr);
      conditionalFormat = dataRange.conditionalFormats.add(
        Excel.ConditionalFormatType.custom
      );

      // Set the font of negative numbers to red.
      conditionalFormat.custom.format.font.color = 'red';
      conditionalFormat.custom.format.fill.color = '#F4B084';
      conditionalFormat.custom.rule.formula = '= NOT(LEFT(E' + (row_start + 2) + ',2) ="0x")';


      for (let row_index = row_start + 5; row_index < row_start + 9; ++row_index) {
        rangestr = "B" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "用格式“[MSB:LSB]”的字符串标识数位字段",
          showPrompt: true,
          title: "Field Range"
        };

        let conditionalFormat = dataRange.conditionalFormats.add(
          Excel.ConditionalFormatType.custom
        );

        // Set the font of negative numbers to red.
        conditionalFormat.custom.format.font.color = 'red';
        conditionalFormat.custom.format.fill.color = '#F4B084';
        conditionalFormat.custom.rule.formula = '= IF(AND( (LEFT(B' + row_index + ',1) ="["), (RIGHT(B'
          + row_index + ',1) = "]" ) ),IF( AND( LEN(B' + row_index + ')>4, ISNUMBER(FIND(":",B' + row_index + ') ) ),FALSE,TRUE), TRUE) ';

        //'= IF(AND( (LEFT(B8,1) ="["), (RIGHT(B8,1) = "]" ),FIND(":",B8,4) ),FALSE, TRUE)  '

        rangestr = "C" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "标识field的字符串。字段name在寄存器范围内必须是唯一的，且为数位段名称首字母大写的缩写。为了能生成头文件，该名称必须是ANSI C标识符",
          showPrompt: true,
          title: "name"
        };

        rangestr = "D" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "定义为字段的默认访问权限",
          showPrompt: true,
          title: "access"
        };

        listRule = {
          list: {
            source: "R,W,RW,RW1C,RW0C,RC,RS,RWT,WT,W1C,W1S,RWRL,Res",
            inCellDropDown: true
          }
        };
        dataRange.dataValidation.rule = listRule;

        rangestr = "E" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "定义复位后所有寄存器的默认值",
          showPrompt: true,
          title: "defaultValue"
        };

        textLenBetwRule = {
          textLength: {
            formula1: 3,
            formula2: 10,
            operator: Excel.DataValidationOperator.between
          }
        };
        dataRange.dataValidation.rule = textLenBetwRule;

        conditionalFormat = dataRange.conditionalFormats.add(
          Excel.ConditionalFormatType.custom
        );

        // Set the font of negative numbers to red.
        conditionalFormat.custom.format.font.color = 'red';
        conditionalFormat.custom.format.fill.color = '#F4B084';
        conditionalFormat.custom.rule.formula = '= NOT(LEFT(E' + row_index + ',2) ="0x")';

        rangestr = "F" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "定义写入寄存器的数值的约束属性。\
          enumerated: 只能写入枚举元素列表中列出的值。\
          range: 只能写入range元素指定的minimum和maximum之间的值",
          showPrompt: true,
          title: "writeConstraint"
        };

        listRule = {
          list: {
            source: "enumerated,range",
            inCellDropDown: true
          }
        };
        dataRange.dataValidation.rule = listRule;

        rangestr = "G" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "指定要写入的最小数字",
          showPrompt: true,
          title: "minimum"
        };

        rangestr = "H" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "指定要写入的最大数字",
          showPrompt: true,
          title: "maximum"
        };

        rangestr = "I" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "描述值语义的字符串",
          showPrompt: true,
          title: "name"
        };

        rangestr = "J" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "用十进制定义位字段的常量",
          showPrompt: true,
          title: "value"
        };

        rangestr = "K" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "描述值的额外说明字符串",
          showPrompt: true,
          title: "description"
        };

        rangestr = "L" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "描述位字段信息的字符串",
          showPrompt: true,
          title: "description"
        };


        rangestr = "B" + row_index + ':E' + row_index;
        dataRange = sheet.getRange(rangestr);
        conditionalFormat = dataRange.conditionalFormats.add(
          Excel.ConditionalFormatType.custom
        );
        conditionalFormat.priority = 0;
        conditionalFormat.stopIfTrue = true;
        // Set the font of negative numbers to red.
        conditionalFormat.custom.format.borders.bottom.color = "red";
        conditionalFormat.custom.format.borders.top.color = "red";
        conditionalFormat.custom.format.borders.left.color = "red";
        conditionalFormat.custom.format.borders.right.color = "red";
        conditionalFormat.custom.format.fill.color = '#FFC3C3';

        conditionalFormat.custom.rule.formula = '=ISBLANK(B' + row_index + ')';
      }


      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      // await context.sync();
      // console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function addCluster() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Create the headers and format them to stand out.
      let headers = [
        ["cluster"]
      ];
      range.load("rowIndex");
      const usedRange = sheet.getUsedRange();
      usedRange.load('columnCount');
      await context.sync();
      const row_start = range.rowIndex + 1;
      let row = row_start;

      const columnCount = 14 > usedRange.columnCount ? 14 : usedRange.columnCount;
      //const columnCount = 14;

      const columnstr = String.fromCharCode(65 + columnCount - 1);
      let rangestr = "A" + row + ":" + columnstr + row;
      let headerRange = sheet.getRange(rangestr);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);
      headerRange.insert(Excel.InsertShiftDirection.down);

      rangestr = "A" + row_start + ":L" + (row_start + 5);
      const bodyRange = sheet.getRange(rangestr);
      bodyRange.conditionalFormats.clearAll();
      bodyRange.dataValidation.clear();
      bodyRange.format.borders.getItem('InsideHorizontal').style = 'Continuous';
      bodyRange.format.borders.getItem('InsideVertical').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeRight').style = 'Continuous';
      bodyRange.format.borders.getItem('EdgeTop').style = 'Continuous';

      bodyRange.format.borders.getItem('EdgeTop').weight = 'Medium';
      bodyRange.format.borders.getItem('EdgeBottom').weight = 'Medium';
      bodyRange.format.borders.getItem('EdgeLeft').weight = 'Medium';
      bodyRange.format.borders.getItem('EdgeRight').weight = 'Medium';


      rangestr = 'A' + row
      headerRange = sheet.getRange(rangestr);
      headerRange.values = headers;

      rangestr = "A" + row + ":L" + row;
      headerRange = sheet.getRange(rangestr);
      headerRange.merge();

      headerRange.format.fill.color = "#D0CECE";
      headerRange.format.font.color = "Black";
      headerRange.format.font.bold = true;
      headerRange.format.horizontalAlignment = 'Left';

      headers = [["name", "addressOffset", "alternateCluster", "alternateGroupName", "headerStructName", "dim", "dimIncrement", "dimName", "description", "", "", ""]
      ];
      row += 1;
      rangestr = "A" + row + ":L" + row;
      headerRange = sheet.getRange(rangestr);
      headerRange.values = headers;
      headerRange.format.fill.color = "#D0CECE";
      headerRange.format.font.color = "Black";
      headerRange.format.font.bold = true;


      for (let row_index = row_start + 2; row_index < row_start + 5; ++row_index) {
        rangestr = "A" + row_index;
        let dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "标识cluster的字符串",
          showPrompt: true,
          title: "name"
        };
        let textLenBetwRule = {
          textLength: {
            formula1: 2,
            formula2: 16,
            operator: Excel.DataValidationOperator.between
          }
        };
        dataRange.dataValidation.rule = textLenBetwRule;

        rangestr = "B" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "相对于器件的“baseAddress”的cluster地址偏移。0x开始的十六进制字串",
          showPrompt: true,
          title: "name"
        };
        textLenBetwRule = {
          textLength: {
            formula1: 4,
            formula2: 6,
            operator: Excel.DataValidationOperator.between
          }
        };
        dataRange.dataValidation.rule = textLenBetwRule;
        let conditionalFormat = dataRange.conditionalFormats.add(
          Excel.ConditionalFormatType.custom
        );

        // Set the font of negative numbers to red.
        conditionalFormat.custom.format.font.color = 'red';
        conditionalFormat.custom.format.fill.color = '#F4B084';
        conditionalFormat.custom.rule.formula = '= NOT(LEFT(E' + row_index + ',2) ="0x")';

        rangestr = "C" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "如果当前簇是一个可替换簇的描述，指定原始簇的名称",
          showPrompt: true,
          title: "alternate"
        };



        rangestr = "E" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "指定在头文件中创建的结构体的名称。如果未指定，则使用cluster的name",
          showPrompt: true,
          title: "headerStructName"
        };

        rangestr = "F" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "定义cluster数组",
          showPrompt: true,
          title: "dim"
        };

        let wholeRule = {
          wholeNumber: {
            formula1: 2,
            formula2: 16,
            operator: Excel.DataValidationOperator.between
          }
        };
        dataRange.dataValidation.rule = wholeRule;

        rangestr = "G" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "指定两个相邻元素之间的地址增量（以字节为单位",
          showPrompt: true,
          title: "dimIncrement"
        };

        rangestr = "H" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "指定C类型结构的名称。如果未定义，则使用寄存器“name",
          showPrompt: true,
          title: "dimName"
        };

        textLenBetwRule = {
          textLength: {
            formula1: 2,
            formula2: 8,
            operator: Excel.DataValidationOperator.between
          }
        };
        dataRange.dataValidation.rule = textLenBetwRule;


        rangestr = "I" + row_index;
        dataRange = sheet.getRange(rangestr);
        dataRange.dataValidation.prompt = {
          message: "描述寄存器簇信息的字符串",
          showPrompt: true,
          title: "description"
        };

        rangestr = "A" + row_index + ':B' + row_index;
        dataRange = sheet.getRange(rangestr);
        conditionalFormat = dataRange.conditionalFormats.add(
          Excel.ConditionalFormatType.custom
        );
        conditionalFormat.priority = 0;
        conditionalFormat.stopIfTrue = true;
        // Set the font of negative numbers to red.
        conditionalFormat.custom.format.borders.bottom.color = "red";
        conditionalFormat.custom.format.borders.top.color = "red";
        conditionalFormat.custom.format.borders.left.color = "red";
        conditionalFormat.custom.format.borders.right.color = "red";
        conditionalFormat.custom.format.fill.color = '#FFC3C3';

        conditionalFormat.custom.rule.formula = '=ISBLANK(B' + row_index + ')';
      }

      headers = [["end cluster"]
      ];
      row = row_start + 5;
      rangestr = "A" + row;
      headerRange = sheet.getRange(rangestr);
      headerRange.values = headers;

      rangestr = "A" + (row_start + 5) + ":L" + (row_start + 5);
      headerRange = sheet.getRange(rangestr);
      headerRange.merge();

      headerRange.format.fill.color = "#D0CECE";
      headerRange.format.font.color = "Black";
      headerRange.format.font.bold = true;
      headerRange.format.horizontalAlignment = 'Left';



      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      // await context.sync();
      // console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
