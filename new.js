/**
 * Copyright by XmT Ltd.
 * http://www.xiaomantou.net
 */
let Excel = require("exceljs");
let filename = "new.xlsx";
let workbook = new Excel.Workbook();

workbook.views = [
    {
        x: 0, y: 0, width: 10000, height: 20000,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
    }
];

let sheet = workbook.addWorksheet("My Sheet");

sheet.addRow(["施工单位", "项目名称", "C30单价（元/m³）", "月份", "当月生产方量(m3)", "当月生产价值(元)", "累计生产方量(m3)", "累计生产价值(元)", "余存方量(m3)", "投资部份", "合同规定当月应收货款（元）", "当月实收货款（元）", "累计实收货款（元）", "合同规定未收货款累计（元）", "所欠货款总计（元）", "备注"]);

workbook.xlsx.writeFile(filename)
    .then(() => {
        console.log("success");
    });