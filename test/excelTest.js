const xlsx = require('xlsx');

const filePath = 'test-data.xlsx';
const outputFile = 'test-data-result.xlsx';

const workbook = xlsx.readFile(filePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

function extractTextInQuotes(str) {
  const regex = /"([^"]+)"/g;
  let matches = [];
  let match;
  while ((match = regex.exec(str)) !== null) {
    matches.push(match[1]);
  }
  return matches;
}

module.exports = {
  '@tags': ['excel-ui'],

  'Thực hiện automation từ mô tả trong Excel': async function (browser) {
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const actions = extractTextInQuotes(row['Mô tả']);
      const expectedText = row['Kết quả'];

      try {
        await browser
          .url('https://actvn.edu.vn')
          .waitForElementVisible('body', 5000)
          .useXpath();

        for (let j = 0; j < actions.length; j++) {
          const actionText = actions[j];

          const xpath = `//*[text()[contains(., "${actionText}")]]`;

          await browser.waitForElementVisible(xpath, 5000);

          if (j === 0) {
            // Hover
            await browser.moveToElement(xpath, 5, 5);
          } else {
            // Click
            await browser.click(xpath);
            await browser.pause(2000);
          }
        }

        // Kiểm tra nội dung
        await browser.useCss();
        await browser.assert.textContains('body', expectedText);

        rows[i]['Kết quả kiểm thử'] = 'PASS';
      } catch (error) {
        rows[i]['Kết quả kiểm thử'] = 'FAIL';
        console.error(`❌ Lỗi dòng ${i + 2}:`, error);
      }
    }

    // Ghi kết quả vào file mới
    const resultSheet = xlsx.utils.json_to_sheet(rows);
    const resultBook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(resultBook, resultSheet, 'Sheet1');
    xlsx.writeFile(resultBook, outputFile);

    browser.end();
  }
};
