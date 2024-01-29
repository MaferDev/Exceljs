"use strict";

const Excel = require("exceljs");
const data = require("./data");

let workbook = new Excel.Workbook();

//*****************************************

let worksheet = workbook.addWorksheet("Tallas", {
  views: [{ state: "frozen", ySplit: 1 }],
});
worksheet.columns = [
  { header: "FieldValueId", key: "FieldValueId", width: 20 },
  { header: "Value", key: "Value", width: 30 },
  { header: "IsActive", key: "IsActive", width: 20 },
  { header: "Position", key: "Position", width: 15 },
];

worksheet.getRow(1).font = {
  size: 11,
  family: 2,
  name: "Calibri",
  bold: true,
  color: { argb: "FFFFFF" },
};

worksheet.getRow(1).fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFC0000" },
};
for (const iterator of data.variedades) {
  worksheet.addRow(iterator);
}

// Keep in mind that reading and writing is promise based.
workbook.xlsx.writeFile("variedades.xlsx");
