/**
 * Copyright by XmT Ltd.
 * http://www.xiaomantou.net
 */
Date.prototype.Format = function (fmt) { //author: meizz
    let o = {
        "M+": this.getMonth() + 1, //月份
        "d+": this.getDate(), //日
        "h+": this.getHours(), //小时
        "m+": this.getMinutes(), //分
        "s+": this.getSeconds(), //秒
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度
        "S": this.getMilliseconds() //毫秒
    };
    if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (let k in o)
        if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    return fmt;
};

let Excel = require("exceljs");
let filename = "合同.xlsx";
let newFilename = "new.xlsx";
let oldWorkbook = new Excel.Workbook();

let newWorkbook = new Excel.Workbook();
let newSheet = newWorkbook.addWorksheet("My Sheet");
newSheet.addRow(["施工单位", "项目名称", "C30单价（元/m³）", "供货期间", "当月生产方量(m3)", "当月生产价值(元)", "累计生产方量(m3)", "累计生产价值(元)", "余存方量(m3)", "投资部份", "合同规定当月应收货款（元）", "当月实收货款（元）", "累计实收货款（元）", "合同规定未收货款累计（元）", "所欠货款总计（元）", "备注"]);
newWorkbook.views = [
    {
        x: 0, y: 0, width: 10000, height: 20000,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
    }
];

oldWorkbook.xlsx.readFile(filename)
    .then(function () {
        oldWorkbook.eachSheet((worksheet) => {
            if (worksheet.name !== "总" || worksheet.name !== "Sheet1") {
                handleSheet(worksheet);
            }
        });
        newWorkbook.xlsx.writeFile(newFilename)
            .then(() => {
                console.log("write success");
            })
            .catch((err) => {
                console.error(err)
            })
    })
    .catch(function (err) {
        console.error(err)
    });

function handleSheet(worksheet) {
    let sheetName = worksheet.name;
    let builder = worksheet.getCell("B2").value;
    let payMethod = worksheet.getCell("J1").value;

    if (payMethod && payMethod.richText !== undefined) {
        let newPayMethod = "";
        payMethod.richText.forEach((textObj) => {
            newPayMethod += textObj.text;
        });
        payMethod = newPayMethod;
    }

    let projectNameCol = worksheet.getColumn('A');
    let columnRowNumbers = [];

    let originalDate = "";
    let finalDate = "";

    projectNameCol.eachCell((cell, rowNumber) => {
        if (cell.value !== null) {
            columnRowNumbers.push(rowNumber)
        }
        if (cell.value instanceof Date && !originalDate) {
            originalDate = (new Date(cell.value).Format("yyyy年MM月dd"));
        }
    });

    let row = worksheet.getRow(columnRowNumbers[columnRowNumbers.length - 1]);
    let newRow = [builder, sheetName, payMethod, ""];

    row.eachCell((cell, colNumber) => {
        if (colNumber === 1) {
            finalDate = (new Date(cell.value).Format("yyyy年MM月dd"));
        } else if (cell.value && cell.value.result !== undefined) {
            newRow.push(cell.value.result)
        } else {
            newRow.push(cell.value)
        }
    });
    newRow[3] = originalDate + "-" + finalDate;
    newSheet.addRow(newRow)
}