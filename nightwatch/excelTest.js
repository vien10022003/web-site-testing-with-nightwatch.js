const xlsx = require("xlsx");

const filePath = "test/ACTVN_TestCases (1).xlsx";
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
        matchedText: match[0],  
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

let actionSrcText;

async function runTestCase(actions, expectedList, browser) {
  for (let i = 0; i < actions.length; i++) {
    const actionJson = actions[i];
    const actionType = actionJson.action;
    actionSrcText = actionJson.matchedText;
    await switchAndRunAction(actionJson, actionType, browser);
  }

  for (let i = 0; i < expectedList.length; i++) {
    const actionJson = expectedList[i];
    const actionType = actionJson.action;
    await switchAndRunAction(actionJson, actionType, browser);
  }
}

// Hàm hỗ trợ: Tạo XPath so sánh text có dấu (tiếng Việt)
function createTextMatchXpath(inputText) {
  const upperChars =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ";
  const lowerChars =
    "abcdefghijklmnopqrstuvwxyzđáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
  const lowerText = inputText.toLowerCase();
  return `//*[translate(normalize-space(text()), '${upperChars}', '${lowerChars}') = "${lowerText}"]`;
}

// Hàm hỗ trợ: Tạo XPath so sánh text có dấu (tiếng Việt)
function createTextMatchXpathContain(inputText) {
  const upperChars =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ";
  const lowerChars =
    "abcdefghijklmnopqrstuvwxyzđáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
  const lowerText = inputText.toLowerCase();
  return `//*[contains(translate(normalize-space(text()), '${upperChars}', '${lowerChars}'), "${lowerText}")]`;
}

async function switchAndRunAction(action, type, browser) {
  console.log(`➡️ Thực hiện action: ${type} với dữ liệu:`, action);

  await browser.useXpath();

  switch (type) {
    case "hover": {
      actionTargetText = action.targetText;
      const xpath = createTextMatchXpath(actionTargetText);
      console.log(`  - Hover đến phần tử có text: "${actionTargetText}" (XPath: ${xpath})`);
      await browser.waitForElementPresent("xpath", xpath, 5000);
      await browser.moveToElement(xpath, 5, 5);
      break;
    }

    case "click": {
      actionTargetText = action.targetText;
      const xpath = createTextMatchXpath(actionTargetText);
      console.log(`  - Click vào phần tử có text: "${actionTargetText}" (XPath: ${xpath})`);
      await browser.waitForElementPresent("xpath", xpath, 5000);
      await browser.click(xpath);
      await browser.pause(1000);
      break;
    }

    case "dropdown_click": {
      actionTargetText = `Parent: ${action.parentText}, Child: ${action.childText}`;
      const parentXpath = createTextMatchXpath(action.parentText);
      const childXpath = createTextMatchXpath(action.childText);
      console.log(`  - Hover vào "${action.parentText}" rồi click vào "${action.childText}"`);
      await browser.waitForElementPresent("xpath", parentXpath, 5000);
      await browser.moveToElement(parentXpath, 5, 5);
      await browser.pause(1000);
      await browser.waitForElementPresent("xpath", childXpath, 5000);
      await browser.click(childXpath);
      break;
    }

    case "scroll": {
      actionTargetText = `scroll-${action.direction}`;
      const direction = action.direction === "down" ? 1000 : -1000;
      console.log(`  - Cuộn trang theo chiều: ${action.direction}`);
      await browser.execute(`window.scrollBy(0, ${direction})`);
      await browser.pause(500);
      break;
    }

    case "click_input_by_label": {
      actionTargetText = action.label;
      const labelXpath = `${createTextMatchXpath(action.label)}/following::input[1]`;
      console.log(`  - Click vào ô input gần label: "${actionTargetText}"`);
      await browser.waitForElementPresent("xpath", labelXpath, 5000);
      await browser.click(labelXpath);
      break;
    }

    case "type": {
      actionTargetText = `type: ${action.value}`;
      console.log(`  - Gõ nội dung: "${action.value}" vào ô đã focus`);
      await browser.setValue("xpath", "//input | //textarea", action.value);
      await browser.pause(500);
      break;
    }

    case "press_key": {
      actionTargetText = `press_key: ${action.key}`;
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
      actionTargetText = `Drag: ${action.sourceText} → ${action.targetText}`;
      const sourceXpath = createTextMatchXpath(action.sourceText);
      const targetXpath = createTextMatchXpath(action.targetText);
      console.log(`  - Kéo "${action.sourceText}" và thả vào "${action.targetText}"`);
      await browser.waitForElementPresent("xpath", sourceXpath, 5000);
      await browser.waitForElementPresent("xpath", targetXpath, 5000);
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
      actionTargetText = action.dropdownText;
      const dropdownXpath = `${createTextMatchXpath(action.dropdownText)}/following::select[1]`;
      console.log(`  - Chọn "${action.value}" từ dropdown gần "${action.dropdownText}"`);
      await browser.waitForElementPresent("xpath", dropdownXpath, 5000);
      await browser.setValue(dropdownXpath, action.value);
      break;
    }

    case "wait": {
      actionTargetText = `wait ${action.seconds}s`;
      console.log(`  - Chờ trong ${action.seconds} giây`);
      await browser.pause(action.seconds * 1000);
      break;
    }

    case "check_count": {
      actionTargetText = action.text;
      const xpath = createTextMatchXpath(action.text);
      console.log(`  - Kiểm tra số lượng phần tử chứa text "${action.text}" là ${action.expectedCount}`);
      await browser.elements("xpath", xpath, function (res) {
        this.assert.equal(res.value.length, action.expectedCount);
      });
      break;
    }

    case "check_visible": {
      actionTargetText = action.text;
      const xpath = createTextMatchXpathContain(action.text);
      console.log(`  - Kiểm tra phần tử chứa text "${action.text}" hiển thị trên giao diện`);
      await browser.waitForElementPresent("xpath", xpath, 3000);
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
    for (let i = 0; i < rows.length; i++) {
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
        await browser.waitForElementPresent("body", 3000);
        // console.log(`✅ Trang đã sẵn sàng`);
        await browser.useXpath(); // chuyển lại XPATH nếu cần sau đó
        // console.log(`🔄 Chuyển sang chế độ XPATH`);

        await runTestCase(actions, expectedList, browser);

        // // Kiểm tra kết quả mong đợi (nếu có)
        // if (expectedText) {
        //   await browser.useCss();
        //   await browser.assert.textContains("body", expectedText);
        // }

        rows[i]["Trạng thái (Pass/Fail)"] = "PASS";
        rows[i]["Kết quả thực tế (sau khi chạy script)"] = rows[i]["Kết quả mong đợi (theo code)"];
        console.log(`✅ PASS dòng ${i + 2}`);
      } catch (error) {
        await browser.waitForElementPresent("xpath", "//body", 60000);
        rows[i]["Trạng thái (Pass/Fail)"] = `FAIL`;
        rows[i]["Kết quả thực tế (sau khi chạy script)"] = `Không thể thực hiện hành động "${actionSrcText}"`;
        console.log(`❌ Lỗi dòng ${i + 2}:`, error.message || error);
      }
    }
await browser.url(browser.launch_url);
await browser.waitForElementPresent("body", 3000);


    // Ghi kết quả vào file
    const resultSheet = xlsx.utils.json_to_sheet(rows);
    const resultBook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(resultBook, resultSheet, "Sheet1");
    xlsx.writeFile(resultBook, outputFile);

    browser.end();
  },
};
