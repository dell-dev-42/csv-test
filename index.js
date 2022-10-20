const csv = require("csvtojson");
const Excel = require("exceljs");
const csvWorksnapsFilePath = "./worksnaps.csv";
const csvUpworkFilePath = "./upwork.csv";
const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("ExampleSheet");
const lodash = require('lodash');

function toMS(str) {
  if (!str.includes(":")) return parseFloat(str);
  const [mins, secms] = str.split(":");
  const [sec] = secms.split(".");
  return (+mins * 60 + +sec) / 60;
}

(async () => {
  const jsonWorksnaps = await csv().fromFile(csvWorksnapsFilePath);
  const jsonUpwork = await csv().fromFile(csvUpworkFilePath);

  const worksnapsData = jsonWorksnaps
    .filter(
      (key) => key.field7 != null && key.field7 !== " " && key.field4 != "Date"
    )
    .map((el) => ({
      date_worksnaps: el.field4,
      time_worksnaps: toMS(el.field7),
    }));

  const upworkData = jsonUpwork.map((el) => ({
    date_upwork: el.Date,
    time_upwork: parseFloat(el.Hours),
  }));

  // add column headers
  worksheet.columns = [
    { header: "Date Worksnaps", key: "date_worksnaps" },
    { header: "Time Worksnaps", key: "time_worksnaps" },
    { header: "Time Upwork", key: "date_upwork" },
    { header: "Time Upwork", key: "time_upwork" },
  ];

  // Add row using key mapping to columns
  worksheet.addRows(lodash.merge(worksnapsData, upworkData));

  worksheet.getCell("B25").value = { formula: "SUM(B2:B24)", result: 7 };
  worksheet.getCell("D25").value = { formula: "SUM(D2:D24)", result: 7 };

  // save workbook to disk
  workbook.xlsx
    .writeFile("report.xlsx")
    .then(() => {
      console.log("File created.");
    })
    .catch((err) => {
      console.log("err", err);
    });
})();
