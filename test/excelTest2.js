const xlsx = require("xlsx");

// const filePath = "test/ACTVN_TestCases (1).xlsx";
const filePath = "test/ACTVN_TestCases - test.xlsx";
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
    {
      type: "check_url_contains",
      regex: /Kiểm tra URL chứa "(.*?)"/g,
      extract: (m) => ({
        action: "check_url_contains",
        text: m[1],
      }),
    },
    {
      type: "check_load_time",
      regex: /Kiểm tra thời gian tải trang < “?(\d+)”? ms/g,
      extract: (m) => ({
        action: "check_load_time",
        maxTime: parseInt(m[1], 10),
      }),
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
    await switchAndRunAction(action, type, browser);
  }

  for (let i = 0; i < expectedList.length; i++) {
    const expected = expectedList[i];
    const { action: expectedtype } = expected;
    await switchAndRunAction(expected, expectedtype, browser);
  }
}

async function generateVariants(text) {
  return [
    text,
    text.toLowerCase(),
    text.toUpperCase(),
    text.charAt(0).toUpperCase() + text.slice(1).toLowerCase(),
    text
      .split(" ")
      .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
      .join(" "),
  ];
}

async function findFlexibleTextAndDo(browser, baseText, callbackFn) {
  const variants = await generateVariants(baseText);

  for (const variant of variants) {
    // const xpath = `//*[contains(normalize-space(string(.)), "${variant}")]`;
    // const xpath = `//*[self::a or self::button or self::span or self::h1 or self::i or self::p or self::h2 or self::label][contains(normalize-space(string(.)), "${variant}")]`;
    const xpath = `//*[self::a or self::button or self::span or self::h1 or self::i or self::p or self::h2 or self::label][contains(string(.), "${variant}")]`;
    // const xpath = `//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"${variant.toLowerCase()}")]`;

    console.log(`➡️ Thử tìm phần tử với text: "${variant}"`);

    try {
      await browser.useXpath();
      

      const fs = require('fs');
      await browser.source(function(result) {
        const logContent = result.value;

        fs.writeFileSync('test/log.html', logContent, { flag: 'w' }); // flag: 'a' là append
      });
      
      await browser.waitForElementPresent("xpath", xpath, 2000);

      // 🔍 Log vị trí phần tử
      await browser.execute(
        function (xpath) {
          const result = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
          );
          const element = result.singleNodeValue;
          if (!element) return null;
          const rect = element.getBoundingClientRect();
          return {
            x: rect.x,
            y: rect.y,
            width: rect.width,
            height: rect.height,
            text: element.innerText || element.textContent,
          };
        },
        [xpath],
        function (result) {
          if (result.value) {
            console.log(
              `📍 Vị trí phần tử: x=${result.value.x.toFixed(
                0
              )}, y=${result.value.y.toFixed(
                0
              )}, width=${result.value.width.toFixed(
                0
              )}, height=${result.value.height.toFixed(
                0
              )}, nội dung: "${result.value.text.trim()}"`
            );
          } else {
            console.warn("⚠️ Không thể lấy thông tin vị trí phần tử.");
          }
        }
      );

      await browser.useXpath();

      await callbackFn(browser, xpath, variant);
      console.log(`✅ Thành công với biến thể: "${variant}"`);
      return true;
    } catch (err) {
      console.warn(
        `⚠️ Không tìm thấy phần tử với biến thể: "${variant}"`,
        err.message || err
      );
      console.warn(`🔄 await browser.waitForElementVisible("body", 3000);`);
      await browser.waitForElementPresent("xpath", "//body", 10000);
    }
  }

  console.warn(
    `❌ Không tìm thấy phần tử khớp với bất kỳ biến thể nào của: "${baseText}"`
  );
  return false;
}

async function switchAndRunAction(action, type, browser) {
  console.log(`➡️ Thực hiện action: ${type} với dữ liệu:`, action);
  await browser.useXpath();

  switch (type) {
    case "hover": {
      const ok = await findFlexibleTextAndDo(
        browser,
        action.targetText,
        async (browser, xpath1) => {
          console.log(`  - Hover đến phần tử có text: "${action.targetText}"`);
          await browser.waitForElementVisible("xpath", xpath1, 5000);
          await browser.moveToElement(xpath1, 5, 5);
        }
      );
      if (!ok)
        throw new Error(
          `Không tìm thấy phần tử để hover: "${action.targetText}"`
        );
      break;
    }

    case "click": {
      const ok = await findFlexibleTextAndDo(
        browser,
        action.targetText,
        async (browser, xpath) => {
          console.log(`  - Click vào phần tử có text: "${action.targetText}"`);
          await browser.waitForElementVisible("xpath", xpath, 5000);
          await browser.click(xpath);
          await browser.pause(1000);
        }
      );
      if (!ok)
        throw new Error(
          `Không tìm thấy phần tử để click: "${action.targetText}"`
        );
      break;
    }

    case "dropdown_click": {
      const okParent = await findFlexibleTextAndDo(
        browser,
        action.parentText,
        async (browser, parentXpath) => {
          const okChild = await findFlexibleTextAndDo(
            browser,
            action.childText,
            async (browser, childXpath) => {
              console.log(
                `  - Hover vào "${action.parentText}" rồi click vào "${action.childText}"`
              );
              await browser.waitForElementVisible(parentXpath, 5000);
              await browser.moveToElement(parentXpath, 5, 5);
              await browser.pause(1000);
              await browser.waitForElementVisible(childXpath, 5000);
              await browser.click(childXpath);
            }
          );
          if (!okChild)
            throw new Error(
              `Không tìm thấy phần tử con: "${action.childText}"`
            );
        }
      );
      if (!okParent)
        throw new Error(`Không tìm thấy phần tử cha: "${action.parentText}"`);
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
      const ok = await findFlexibleTextAndDo(
        browser,
        action.label,
        async (browser, labelXpath) => {
          const inputXpath = `${labelXpath}/following::input[1]`;
          console.log(`  - Click vào ô input gần label: "${action.label}"`);
          await browser.waitForElementVisible(inputXpath, 5000);
          await browser.click(inputXpath);
        }
      );
      if (!ok)
        throw new Error(
          `Không tìm thấy label để click input: "${action.label}"`
        );
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
      const okSource = await findFlexibleTextAndDo(
        browser,
        action.sourceText,
        async (browser, sourceXpath) => {
          const okTarget = await findFlexibleTextAndDo(
            browser,
            action.targetText,
            async (browser, targetXpath) => {
              console.log(
                `  - Kéo phần tử "${action.sourceText}" và thả vào "${action.targetText}"`
              );
              await browser.waitForElementVisible(sourceXpath, 5000);
              await browser.waitForElementVisible(targetXpath, 5000);
              await browser.perform((done) => {
                browser
                  .moveToElement(sourceXpath, 5, 5)
                  .mouseButtonDown(0)
                  .moveToElement(targetXpath, 5, 5)
                  .mouseButtonUp(0);
                done();
              });
            }
          );
          if (!okTarget)
            throw new Error(
              `Không tìm thấy phần tử đích: "${action.targetText}"`
            );
        }
      );
      if (!okSource)
        throw new Error(`Không tìm thấy phần tử nguồn: "${action.sourceText}"`);
      break;
    }

    case "select_dropdown": {
      const ok = await findFlexibleTextAndDo(
        browser,
        action.dropdownText,
        async (browser, dropdownXpath) => {
          const selectXpath = `${dropdownXpath}/following::select[1]`;
          console.log(
            `  - Chọn "${action.value}" từ dropdown gần "${action.dropdownText}"`
          );
          await browser.waitForElementVisible(selectXpath, 5000);
          await browser.setValue(selectXpath, action.value);
        }
      );
      if (!ok)
        throw new Error(
          `Không tìm thấy dropdown label: "${action.dropdownText}"`
        );
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
      console.log(
        `  - Kiểm tra số lượng phần tử chứa text "${action.text}" là ${action.expectedCount}`
      );
      await browser.elements("xpath", xpath, function (res) {
        this.assert.equal(res.value.length, action.expectedCount);
      });
      break;
    }

    case "check_visible": {
      const ok = await findFlexibleTextAndDo(
        browser,
        action.text,
        async (browser, xpath) => {
          const fs = require("fs");
          await browser.source(function (result) {
            const logContent = result.value;
            fs.writeFileSync("test/log.html", logContent, { flag: "w" });
          });

          console.log(
            `  - Kiểm tra phần tử chứa text "${action.text}" hiển thị trên giao diện`
          );
          await browser.waitForElementPresent("xpath", xpath, 3000);
        }
      );
      if (!ok)
        throw new Error(`Không tìm thấy phần tử hiển thị: "${action.text}"`);
      break;
    }

    case "check_url_contains": {
      console.log(`  - Kiểm tra URL chứa: "${action.text}"`);
      await browser.url(function (result) {
        const currentUrl = result.value;
        console.log(`    URL hiện tại: ${currentUrl}`);
        if (!currentUrl.includes(action.text)) {
          throw new Error(`URL hiện tại không chứa "${action.text}"`);
        }
      });
      break;
    }

    case "check_load_time": {
      console.log(
        `  - Đo thời gian tải lại trang, yêu cầu < ${action.maxTime} ms`
      );
      const startTime = Date.now();
      await browser.refresh();
      await browser.waitForElementPresent("xpath", "//body", 60000);
      const endTime = Date.now();
      const loadTime = endTime - startTime;
      console.log(`    ⏱ Thời gian tải: ${loadTime} ms`);
      if (loadTime >= action.maxTime) {
        throw new Error(
          `Thời gian tải trang quá lâu: ${loadTime} ms (yêu cầu < ${action.maxTime} ms)`
        );
      }
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
        await browser.useCss();
        console.log(`✅ url 1`);
        await browser.useXpath(); // chuyển lại XPATH nếu cần sau đó
        console.log(`✅ url 2`);
        await browser.url(browser.launch_url);
        console.log(`✅ url 3`);
        // console.log(`🌐 Mở trang: ${browser.launch_url}`);
        await browser.pause(1000); // chờ thêm 1 giây trước khi kiểm tra
        console.log(`✅ url 4`);
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
        console.log(`✅ PASS dòng ${i + 2}:`);
      } catch (error) {
        const message = error?.message || String(error);
        rows[i]["Kết quả kiểm thử"] = `FAIL: ${message}`;
        // rows[i]["Kết quả kiểm thử"] = "FAIL";
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
