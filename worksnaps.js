const fs = require("fs");
const csvFilePath = "./worksnaps.csv";
const csv = require("csvtojson");
const ExcelJS = require("exceljs");
const { time } = require("console");
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("MANOWAR");

workbook.creator = "Vasjan";
workbook.lastModifiedBy = "Bot";
workbook.created = new Date(2022, 10, 18);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2022, 10, 18);

workbook.calcProperties.fullCalcOnLoad = true;

(async () => {
  const jsonArray = await csv().fromFile(csvFilePath);

  const tests = jsonArray
    .filter((key) => key.field7 != null && key.field7 !== " ")
    .map((el) => ({
      date: el.field4,
      time: el.field7,
    }));

  console.log("test :>> ", tests);

  // tests.forEach((element) => {
  //   worksheet.insertRow(element.date);
  //   worksheet.insertRow(element.time);
  // });
})();
