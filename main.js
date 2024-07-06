const ExcelJS = require("exceljs");

const workbook = new ExcelJS.Workbook();
workbook.creator = "Me";
workbook.lastModifiedBy = "Her";
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);

const sheet = workbook.addWorksheet("Test");
sheet.columns = [
  {
    header: "Registration",
    key: "registration",
    width: 15,
  },
];

const rows = sheet.getRows(2, 4);
console.log("rows", rows);

for (const row of rows) {
  const cell = row.getCell("registration");
  cell.dataValidation = {
    type: "list",
    allowBlank: true,
    formulae: ["AAA", "BBB"],
  };
}

workbook.xlsx.writeFile("test.xlsx");
