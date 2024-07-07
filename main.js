const ExcelJS = require("exceljs");

const workbook = new ExcelJS.Workbook();
workbook.creator = "Me";
workbook.lastModifiedBy = "Her";
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);

const vehListSheet = workbook.addWorksheet("VehList", { state: "veryHidden" });
vehListSheet.orderNo = 2;
vehListSheet.columns = [
  {
    header: "Registration",
    key: "registration",
    width: 15,
  },
];

vehListSheet.addRows([
  { registration: "AAA-111" },
  { registration: "AAA-222" },
  { registration: "AAA-333" },
  { registration: "AAA-444" },
  { registration: "AAA-555" },
  { registration: "AAA-666" },
  { registration: "AAA-777" },
  { registration: "AAA-888" },
  { registration: "AAA-999" },
  { registration: "AAA-000" },
]);

const sheet = workbook.addWorksheet("Test");
sheet.orderNo = 1;
sheet.columns = [
  {
    header: "Registration",
    key: "registration",
    width: 15,
  },
  {
    header: "Plan name",
    key: "planName",
    width: 15,
  },
];

const rows = sheet.getRows(2, 4);

for (const row of rows) {
  const cell = row.getCell("registration");
  cell.dataValidation = {
    type: "list",
    allowBlank: false,
    formulae: [`${vehListSheet.name}!$A$2:$A$${vehListSheet.rowCount}`],
    error: "Please select registration from the list",
    errorTitle: "Invalid registration",
    errorStyle: "stop",
    showErrorMessage: true,
  };
}

workbook.xlsx.writeFile("output/test.xlsx");
