const xlsx = require("xlsx");
const { parseActionsInOrder } = require("./utils/textUtils");
const { runTestCase } = require("./testRunner/testCaseRunner");

const filePath = "nightwatch/ACTVN_TestCases (1).xlsx";
// const filePath = "nightwatch/ACTVN_TestCases.xlsx";
// const filePath = "nightwatch/ACTVN_TestCases test.xlsx";
const outputFile = "nightwatch/test-data-result.xlsx";

const workbook = xlsx.readFile(filePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

console.log("thiết lập xong:", filePath);

// var actionSrcText;
global.actionSrcText = "";

module.exports = {
  "@tags": ["excel-ui"],

  "Thực hiện automation từ mô tả trong Excel": async function (browser) {
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const actions = parseActionsInOrder(row["Bước thực hiện (theo code)"]);
      const expectedList = parseActionsInOrder(
        row["Kết quả mong đợi (theo code)"]
      );
      console.log(`🔍 Dòng ${i + 2}:`, actions);

      try {
        await browser.url(browser.launch_url);
        await browser.pause(1000);
        await browser.useCss();
        await browser.waitForElementPresent("body", 3000);
        await browser.useXpath();

        await runTestCase(actions, expectedList, browser);

        rows[i]["Trạng thái (Pass/Fail)"] = "PASS";
        rows[i]["Kết quả thực tế (sau khi chạy script)"] =
          rows[i]["Kết quả mong đợi (theo code)"];
        console.log(`✅ PASS dòng ${i + 2}`);
      } catch (error) {
        await browser.waitForElementPresent("xpath", "//body", 10000);
        rows[i]["Trạng thái (Pass/Fail)"] = `FAIL`;
        rows[i][
          "Kết quả thực tế (sau khi chạy script)"
        ] = `Không thể thực hiện hành động "${global.actionSrcText}"`;
        console.log(`❌ Lỗi dòng ${i + 2}:`, error.message || error);
      }
    }

    // Ghi kết quả vào file
    const resultSheet = xlsx.utils.json_to_sheet(rows);
    const resultBook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(resultBook, resultSheet, "Sheet1");
    xlsx.writeFile(resultBook, outputFile);
    console.warn(`✅✅✅✅✅ Kết quả đã được ghi vào file: ${outputFile}`);
    browser.end();
  },
};
