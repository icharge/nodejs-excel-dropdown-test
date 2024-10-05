const ExcelJS = require("exceljs");

const workbook = new ExcelJS.Workbook();
workbook.creator = "Me";
workbook.lastModifiedBy = "Her";
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);

const typeListSheet = workbook.addWorksheet("TypeList", {
  state: "veryHidden",
});
typeListSheet.orderNo = 2;
typeListSheet.columns = [
  {
    header: "Type",
    key: "type",
  },
];

typeListSheet.addRows([{ type: "Basic" }, { type: "Advance" }]);

const vehListSheet = workbook.addWorksheet("VehList", { state: "veryHidden" });
vehListSheet.orderNo = 3;
vehListSheet.columns = [
  {
    header: "Registration",
    key: "vehReg",
    width: 15,
  },
  {
    header: "Fleet ID",
    key: "fleetId",
  },
];

vehListSheet.addRows([
  { vehReg: "AAA-111" },
  { vehReg: "AAA-222" },
  { vehReg: "AAA-333" },
  { vehReg: "AAA-444" },
  { vehReg: "AAA-555" },
  { vehReg: "AAA-666" },
  { vehReg: "AAA-777" },
  { vehReg: "AAA-888" },
  { vehReg: "AAA-999" },
  { vehReg: "AAA-000" },
]);

const sheet = workbook.addWorksheet("Test");
sheet.orderNo = 1;
sheet.columns = [
  {
    header: "Type",
    key: "type",
  },
  {
    header: "Registration",
    key: "vehReg",
    width: 15,
  },
  {
    header: "TripName",
    key: "etaPlanDesc",
    width: 20,
  },
  {
    header: "Sequence",
    key: "order",
    width: 10,
  },
  {
    header: "WaypointName",
    key: "locationName",
    width: 20,
  },
  {
    header: "PlanArrivalDate",
    key: "arriveDate",
  },
  {
    header: "PlanArrivalTime",
    key: "arriveTime",
  },
  {
    header: "PlanDepartureDate",
    key: "leaveDate",
  },
  {
    header: "PlanDepartureTime",
    key: "leaveTime",
  },
  {
    header: "TimeInZone",
    key: "timeInZone",
  },
  {
    header: "Alert1",
    key: "alert1",
  },
  {
    header: "Alert2",
    key: "alert2",
  },
  {
    header: "Alert3",
    key: "alert3",
  },
];

const startAt = 2;

sheet.insertRows(startAt, [
  {
    vehReg: "AAA-111",
    planName: "plan 1",
  },
  {
    vehReg: "AAA-222",
    planName: "plan 2",
  },
]);

const rows = sheet.getRows(startAt, 4);

for (const row of rows) {
  const typeCell = row.getCell("type");
  typeCell.dataValidation = {
    type: "list",
    allowBlank: false,
    formulae: [`${typeListSheet.name}!$A$2:$A$${typeListSheet.rowCount}`],
    error: "Please select type from the list",
    errorTitle: "Invalid type",
    errorStyle: "stop",
    showErrorMessage: true,
  };

  const vehRegCell = row.getCell("vehReg");
  vehRegCell.dataValidation = {
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
