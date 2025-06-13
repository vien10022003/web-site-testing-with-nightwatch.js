const xlsx = require("xlsx");

const filePath = "test/ACTVN_TestCases.xlsx";
const outputFile = "test/test-data-result.xlsx";

const workbook = xlsx.readFile(filePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

console.log("thiết lập xong:", filePath);

function parseActionsInOrder(description) {
  console.log(`\n🔍 Phân tích mô tả: "${description}"`);
  console.log(description);
  const patterns = [
    {
      type: "hover",
      regex: /Di chuột đến "(.*?)"/g,
      extract: (m) => ({ action: "hover", targetText: m[1] }),
    },
    {
      type: "click",
      regex: /Bấm vào "(.*?)"(?! từ)/g,
      extract: (m) => ({ action: "click", targetText: m[1] }),
    },
    {
      type: "dropdown_click",
      regex: /Bấm vào "(.*?)" từ "(.*?)" được xổ xuống/g,
      extract: (m) => ({
        action: "dropdown_click",
        childText: m[1],
        parentText: m[2],
      }),
    },
    {
      type: "scroll",
      regex: /Cuộn chuột (xuống|lên)/g,
      extract: (m) => ({
        action: "scroll",
        direction: m[1] === "xuống" ? "down" : "up",
      }),
    },
    {
      type: "click_input",
      regex: /Bấm vào ô "(.*?)"/g,
      extract: (m) => ({ action: "click_input_by_label", label: m[1] }),
    },
    {
      type: "type",
      regex: /Gõ "(.*?)"/g,
      extract: (m) => ({ action: "type", value: m[1] }),
    },
    {
      type: "press_key",
      regex: /Nhấn nút "(Enter|Tab|Esc)"/gi,
      extract: (m) => ({ action: "press_key", key: m[1].toUpperCase() }),
    },
    {
      type: "drag_drop",
      regex: /Kéo "(.*?)" và thả vào "(.*?)"/g,
      extract: (m) => ({
        action: "drag_drop",
        sourceText: m[1],
        targetText: m[2],
      }),
    },
    {
      type: "select_dropdown",
      regex: /Chọn "(.*?)" từ danh sách "(.*?)"/g,
      extract: (m) => ({
        action: "select_dropdown",
        value: m[1],
        dropdownText: m[2],
      }),
    },
    {
      type: "wait_time",
      regex: /Chờ “?(\d+)”? giây/g,
      extract: (m) => ({ action: "wait", seconds: parseInt(m[1], 10) }),
    },
    {
      type: "check_count",
      regex: /Kiểm tra số lượng "(.*?)" là (\d+)/g,
      extract: (m) => ({
        action: "check_count",
        text: m[1],
        expectedCount: parseInt(m[2], 10),
      }),
    },
    {
      type: "check_visible",
      regex: /Kiểm tra thấy "(.*?)"/g,
      extract: (m) => ({ action: "check_visible", text: m[1] }),
    },
  ];

  const results = [];

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.regex.exec(description)) !== null) {
      results.push({
        index: match.index,
        ...pattern.extract(match),
      });
    }
  }

  // Sắp xếp theo thứ tự xuất hiện
  results.sort((a, b) => a.index - b.index);

  console.log(`\n🔍 Phân tích mô tả: `);
  console.log(results.map(({ index, ...rest }) => rest));
  // Loại bỏ trường 'index'
  return results.map(({ index, ...rest }) => rest);
}

async function runTestCase(actions, expectedList, browser) {
  for (let i = 0; i < actions.length; i++) {
    const action = actions[i];
    const { action: type } = action;
    await switchAndRunAction(action, type, browser).catch((error) => {
      console.error(
        `❌ Lỗi khi thực hiện action ${i + 1}:`,
        error.message || error
      );
    });
  }

  for (let i = 0; i < expectedList.length; i++) {
    const expected = expectedList[i];
    const { action: expectedtype } = expected;
    await switchAndRunAction(expected, expectedtype, browser).catch((error) => {
      console.error(
        `❌ Lỗi khi thực hiện expected ${i + 1}:`,
        error.message || error
      );
    });
  }
}
async function switchAndRunAction(action, type, browser) {
  console.log(`➡️ Thực hiện action: ${type} với dữ liệu:`, action);

  await browser.useXpath();

  switch (type) {
    case "hover": {
      const lowerText = action.targetText.toLowerCase();
      const xpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${lowerText}")]`;
      console.log(`  - Hover đến phần tử có text: "${action.targetText}" (XPath: ${xpath})`);
      await browser.waitForElementVisible('xpath', xpath, 5000);
      await browser.moveToElement(xpath, 5, 5);
      break;
    }

    case "click": {
      const lowerText = action.targetText.toLowerCase();
      const xpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${lowerText}")]`;
      console.log(`  - Click vào phần tử có text: "${action.targetText}" (XPath: ${xpath})`);
      await browser.waitForElementVisible('xpath', xpath, 5000);
      await browser.click(xpath);
      await browser.pause(1000);
      break;
    }

    case "dropdown_click": {
      const parentText = action.parentText.toLowerCase();
      const childText = action.childText.toLowerCase();
      const parentXpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${parentText}")]`;
      const childXpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${childText}")]`;
      console.log(`  - Hover vào "${action.parentText}" (XPath: ${parentXpath}) rồi click vào "${action.childText}" (XPath: ${childXpath})`);
      await browser.waitForElementVisible('xpath', parentXpath, 5000);
      await browser.moveToElement(parentXpath, 5, 5);
      await browser.pause(1000);
      await browser.waitForElementVisible('xpath', childXpath, 5000);
      await browser.click(childXpath);
      break;
    }

    case "scroll": {
      const direction = action.direction === "down" ? 1000 : -1000;
      console.log(`  - Cuộn trang theo chiều: ${action.direction}`);
      await browser.execute(`window.scrollBy(0, ${direction})`);
      await browser.pause(500);
      break;
    }

    case "click_input_by_label": {
      const labelText = action.label.toLowerCase();
      const labelXpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${labelText}")]/following::input[1]`;
      console.log(`  - Click vào ô input gần label: "${action.label}" (XPath: ${labelXpath})`);
      await browser.waitForElementVisible('xpath', labelXpath, 5000);
      await browser.click(labelXpath);
      break;
    }

    case "type": {
      console.log(`  - Gõ nội dung: "${action.value}" vào ô đã focus`);
      await browser.setValue("xpath", "//input | //textarea", action.value);
      await browser.pause(500);
      break;
    }

    case "press_key": {
      const keyMap = {
        ENTER: browser.Keys.ENTER,
        TAB: browser.Keys.TAB,
        ESC: browser.Keys.ESCAPE,
      };
      console.log(`  - Nhấn phím: ${action.key}`);
      await browser.keys(keyMap[action.key] || action.key);
      break;
    }

    case "drag_drop": {
      const sourceText = action.sourceText.toLowerCase();
      const targetText = action.targetText.toLowerCase();
      const sourceXpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${sourceText}")]`;
      const targetXpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${targetText}")]`;
      console.log(`  - Kéo phần tử "${action.sourceText}" và thả vào "${action.targetText}"`);
      await browser.waitForElementVisible('xpath', sourceXpath, 5000);
      await browser.waitForElementVisible('xpath', targetXpath, 5000);
      await browser.perform((done) => {
        browser
          .moveToElement(sourceXpath, 5, 5)
          .mouseButtonDown(0)
          .moveToElement(targetXpath, 5, 5)
          .mouseButtonUp(0);
        done();
      });
      break;
    }

    case "select_dropdown": {
      const dropdownText = action.dropdownText.toLowerCase();
      const dropdownXpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${dropdownText}")]/following::select[1]`;
      console.log(`  - Chọn "${action.value}" từ dropdown gần "${action.dropdownText}" (XPath: ${dropdownXpath})`);
      await browser.waitForElementVisible('xpath', dropdownXpath, 5000);
      await browser.setValue(dropdownXpath, action.value);
      break;
    }

    case "wait": {
      console.log(`  - Chờ trong ${action.seconds} giây`);
      await browser.pause(action.seconds * 1000);
      break;
    }

    case "check_count": {
      const lowerText = action.text.toLowerCase();
      const xpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${lowerText}")]`;
      console.log(`  - Kiểm tra số lượng phần tử chứa text "${action.text}" là ${action.expectedCount}`);
      await browser.elements("xpath", xpath, function (res) {
        this.assert.equal(res.value.length, action.expectedCount);
      });
      break;
    }

    case "check_visible": {
      const lowerText = action.text.toLowerCase();
      const xpath = `//*[contains(translate(string(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "${lowerText}")]`;
      console.log(`  - Kiểm tra phần tử chứa text "${action.text}" hiển thị trên giao diện`);

      const fs = require('fs');
      await browser.source(function(result) {
        const logContent = result.value;

        fs.writeFileSync('test/log.html', logContent, { flag: 'w' }); // flag: 'a' là append
      });

      await browser.waitForElementVisible('xpath', xpath, 3000);
      break;  
    }

    default:
      console.warn(`⚠️ Không hỗ trợ action: ${type}`);
  }

  console.log(`✅ Hoàn thành action: ${type}\n`);
}




module.exports = {
  "@tags": ["excel-ui"],

  "Thực hiện automation từ mô tả trong Excel": async function (browser) {
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const actions = parseActionsInOrder(row["Bước thực hiện (theo code)"]);
      const expectedList = parseActionsInOrder(
        row["Kết quả mong đợi (theo code)"]
      );
      console.log(`🔍 Dòng ${i + 2}:`, actions);

      try {
        await browser.url(browser.launch_url);
        // console.log(`🌐 Mở trang: ${browser.launch_url}`);
        await browser.pause(1000); // chờ thêm 1 giây trước khi kiểm tra
        // console.log(`📝 Chờ 1s hoàn tất`);
        await browser.useCss();
        // console.log(`🔄 Chuyển sang chế độ CSS`);
        await browser.waitForElementVisible("body", 3000);
        // console.log(`✅ Trang đã sẵn sàng`);
        await browser.useXpath(); // chuyển lại XPATH nếu cần sau đó
        // console.log(`🔄 Chuyển sang chế độ XPATH`);

        await runTestCase(actions, expectedList, browser);

        // // Kiểm tra kết quả mong đợi (nếu có)
        // if (expectedText) {
        //   await browser.useCss();
        //   await browser.assert.textContains("body", expectedText);
        // }

        rows[i]["Kết quả kiểm thử"] = "PASS";
        console.log(`✅ PASS dòng ${i + 2}:`, error.message || error);
      } catch (error) {
        rows[i]["Kết quả kiểm thử"] = "FAIL";
        console.log(`❌ Lỗi dòng ${i + 2}:`, error.message || error);
      }
    }

    // Ghi kết quả vào file
    const resultSheet = xlsx.utils.json_to_sheet(rows);
    const resultBook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(resultBook, resultSheet, "Sheet1");
    xlsx.writeFile(resultBook, outputFile);

    browser.end();
  },
};
