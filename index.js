const fs = require("fs");
const csv = require("csvtojson");
const Excel = require("exceljs");
const csvWorksnapsFilePath = "./worksnaps.csv";
const csvUpworkFilePath = "./upwork.csv";
const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("ExampleSheet");
const lodash = require("lodash");

function toMS(str) {
  if (!str.includes(":")) return parseFloat(str);
  const [mins, secms] = str.split(":");
  const [sec] = secms.split(".");
  return (+mins * 60 + +sec) / 60;
}

//if exist worksnaps file
if (fs.existsSync(csvWorksnapsFilePath)) {
  (async () => {
    const jsonWorksnaps = await csv().fromFile(csvWorksnapsFilePath);

    const worksnapsData = jsonWorksnaps
      .filter(
        (key) =>
          key.field7 != null && key.field7 !== " " && key.field4 != "Date"
      )
      .map((el) => ({
        date_worksnaps: el.field4,
        time_worksnaps: toMS(el.field7),
      }));

    let data = worksnapsData;

    // add column headers
    worksheet.columns = [
      { header: "Date Worksnaps", key: "date_worksnaps" },
      { header: "Time Worksnaps", key: "time_worksnaps" },
      { header: "Time Upwork", key: "date_upwork" },
      { header: "Time Upwork", key: "time_upwork" },
    ];

    //if exist upwork file
    if (fs.existsSync(csvUpworkFilePath)) {
      const jsonUpwork = await csv().fromFile(csvUpworkFilePath);
      const upworkData = jsonUpwork.map((el) => ({
        date_upwork: el.Date,
        time_upwork: parseFloat(el.Hours),
      }));

      data = lodash.merge(worksnapsData, upworkData);
    }

    // Add row using key mapping to columns
    worksheet.addRows(data);

    //worksnaps total time
    worksheet.getCell("B27").value = { formula: "SUM(B2:B26)" };
    //upwork total time
    worksheet.getCell("D27").value = { formula: "SUM(D2:D26)" };

    worksheet.getCell("A28").value = "Rate:";
    worksheet.getCell("A29").value = "Total:";
    //worksnaps rate
    worksheet.getCell("B28").value = 3;
    //upwork rate
    worksheet.getCell("D28").value = 3.5;

    //worksnaps total
    worksheet.getCell("B29").value = { formula: "(B27*B28)" };
    //upwork total
    worksheet.getCell("D29").value = { formula: "(D27*D28)" };

    //total
    worksheet.getCell("F29").value = { formula: "SUM(B29,D29)" };

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
} else {
  console.log("Зауснь файл в папку");
}
