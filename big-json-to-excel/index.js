const fs = require("fs");
const path = require("path");
const { format } = require("date-fns");

const xl = require("excel4node");
const wb = new xl.Workbook();
const ws = wb.addWorksheet("XXX sheet name");
const tableHeads = [
  "经营组织",
];
const tableValues = [
  "orggName",
];

fs.readFile(path.normalize(".\\excel.txt"), (error, data) => {
  //   console.log(typeof data, JSON.parse(data.toString("utf-8")));
  let headingColumnIndex = 1;
  tableHeads.forEach((heading) => {
    ws.cell(1, headingColumnIndex++).string(heading);
  });
  let rowIndex = 2;
  const json = JSON.parse(data.toString("utf-8"));
  json.forEach((record) => {
    let columnIndex = 1;
    tableValues.forEach((columnName) => {
      let cellData = record[columnName] ? String(record[columnName]) : "";
      if (columnName.toLowerCase().includes("date")) {
        cellData = cellData ? format(new Date(cellData), "yyyy-MM-dd HH:mm:ss") : "";
      }
      ws.cell(rowIndex, columnIndex++).string(cellData);
    });
    rowIndex++;
  });

  wb.write("sales.xlsx");
});

// let headingColumnIndex = 1;
// tableValues.forEach((heading) => {
//   ws.cell(1, headingColumnIndex++).string(heading);
// });
// let rowIndex = 2;
// reader.on("data", (chunk) => {
//   console.log(chunk.toString(), chunk.toJSON());
//   chunk.forEach((record) => {
//     let columnIndex = 1;
//     Object.keys(record).forEach((columnName) => {
//       ws.cell(rowIndex, columnIndex++).string(record[columnName]);
//     });
//     rowIndex++;
//   });
// });

// wb.write("sales.xlsx");
