const {
  createTextMatchXpath,
  createTextMatchXpathContain,
} = require("../utils/textUtils");

async function switchAndRunAction(action, type, browser) {
  console.log(`➡️ Thực hiện action: ${type} với dữ liệu:`, action);

  await browser.useXpath();

  try {
    switch (type) {
      case "click_by_radio": {
        const labelText = action.text;

        const upperChars =
          "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ";
        const lowerChars =
          "abcdefghijklmnopqrstuvwxyzđáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
        const lowerText = labelText.toLowerCase();
        const a = `//label[contains(translate(normalize-space(string(.)), '${upperChars}', '${lowerChars}'), "${lowerText}")]/input[@type="radio"]`;
        console.log(`  - Click radio gần text "${labelText}"`);

        // const fs = require("fs");
        // await browser.source(function (result) {
        //   const logContent = result.value;

        //   fs.writeFileSync("log.html", logContent, { flag: "w" }); // flag: 'a' là append
        // });

        await browser.useXpath().waitForElementPresent(a, 5000);
        await browser.click(a);

        break;
      }

      case "upload_file": {
        const absolutePath = require("path").resolve(
          __dirname + `/../resources/${action.filename}`
        );
        actionTargetText = `upload: ${action.filename} to #${action.elementId}`;
        const selector = `#${action.elementId}`;

        console.log(
          `  - Upload tệp "${action.filename}" vào ô có id: "${action.elementId}"`
        );
        browser.useCss();
        await browser.waitForElementPresent("css selector", selector, 5000);
        await browser.setValue(selector, absolutePath);
        browser.useXpath();
        break;
      }

      case "alert_check_and_accept": {
        actionTargetText = `alert: ${action.expectedText}`;
        console.log(`  - Kiểm tra nội dung alert và nhấn OK`);

        const result = await browser.pause(500).getAlertText();

        const actual = result.trim();
        const expected = action.expectedText.trim();

        console.log(`  - Nội dung alert là: "${actual}"`);

        if (actual != expected) {
          throw new Error(
            `Nội dung alert không khớp: "${actual}" !== "${expected}"`
          );
        }

        await browser.acceptAlert();
        break;
      }

      case "navigate": {
        actionTargetText = `visit: ${action.url}`;
        console.log(`  - Truy cập vào trang: ${action.url}`);
        await browser.waitForElementPresent("xpath", "//body", 10000);
        await browser.url(action.url);
        break;
      }
      case "click_by_id": {
        actionTargetText = `#${action.id}`;
        const selector = `#${action.id}`;
        console.log(`  - Click vào thẻ có id = "${action.id}"`);
        await browser.useCss();
        await browser.waitForElementPresent("css selector", selector, 5000);
        await browser.click("css selector", selector);
        await browser.useXpath();
        break;
      }
      case "hover": {
        actionTargetText = action.targetText;
        const xpath = createTextMatchXpath(actionTargetText);
        console.log(
          `  - Hover đến phần tử có text: "${actionTargetText}" (XPath: ${xpath})`
        );
        await browser.waitForElementPresent("xpath", xpath, 5000);
        await browser.moveToElement(xpath, 5, 5);
        break;
      }

      case "click": {
        actionTargetText = action.targetText;

        // Các thẻ có thể click
        const clickableTags = ["button", "a", "span", "label", "li"];
        const tagConditions = clickableTags
          .map((tag) => `self::${tag}`)
          .join(" or ");

        // const xpath = `//*[self::a or self::button or self::span or self::h1 or self::i or self::p or self::h2 or self::label][contains(string(.), "${variant}")]`;

        const xpath = `//*[${tagConditions}] ${createTextMatchXpath(
          actionTargetText
        ).replace(/^\/\/\*/, "")}`;

        console.log(
          `  - Click vào phần tử có text: "${actionTargetText}" (XPath: ${xpath})`
        );

        await browser.useXpath();
        await browser.waitForElementPresent(xpath, 5000);
        await browser.click(xpath);
        await browser.pause(1000);
        break;
      }

      case "dropdown_click": {
        actionTargetText = `Parent: ${action.parentText}, Child: ${action.childText}`;

        // Tìm element chứa text "Ngày", "Tháng", "Năm", v.v.
        const parentXpath = createTextMatchXpath(action.parentText);

        // Tìm <select> đầu tiên phía sau phần tử chứa text "Ngày"
        const dropdownXpath = `${parentXpath}//parent::select`;

        // Tìm <option> có text là "10" bên trong <select> đó
        const childXpath = `${dropdownXpath}/option[text()="${action.childText}"]`;

        console.log(
          `  - Chọn "${action.childText}" từ dropdown gần "${action.parentText}"`
        );

        // Chờ dropdown xuất hiện và chọn option
        await browser.waitForElementPresent("xpath", dropdownXpath, 5000);
        await browser.click(dropdownXpath); // focus select nếu cần
        await browser.setValue(dropdownXpath, action.childText); // chọn option
        break;
      }

      case "select_dropdown_by_id": {
        const selectXpath = `//select[@id="${action.id}"]`;
        const optionXpath = `${selectXpath}/option[text()="${action.value}"]`;

        console.log(
          `  - Chọn "${action.value}" từ dropdown có id "${action.id}"`
        );

        await browser.useXpath();
        await browser.waitForElementPresent("xpath", selectXpath, 5000);
        await browser.click(selectXpath);
        await browser.click(optionXpath);
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
        const labelXpath = `${createTextMatchXpath(
          action.label
        )}/following::input[1]`;
        console.log(`  - Click vào ô input gần label: "${actionTargetText}"`);
        await browser.waitForElementPresent("xpath", labelXpath, 5000);
        await browser.click(labelXpath);
        break;
      }

      case "type": {
        actionTargetText = `type: ${action.value}`;
        console.log(
          `  - Gõ nội dung: "${action.value}" vào ô đã được focus sẵn`
        );
        // Gán nội dung trực tiếp vào phần tử đang được focus
        await browser.execute(
          function (value) {
            if (
              document.activeElement &&
              (document.activeElement.tagName === "INPUT" ||
                document.activeElement.tagName === "TEXTAREA")
            ) {
              document.activeElement.value = value;
            }
          },
          [action.value]
        );
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
        console.log(
          `  - Kéo "${action.sourceText}" và thả vào "${action.targetText}"`
        );
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

      case "wait": {
        actionTargetText = `wait ${action.seconds}s`;
        console.log(`  - Chờ trong ${action.seconds} giây`);
        await browser.pause(action.seconds * 1000);
        break;
      }

      case "check_count": {
        actionTargetText = action.text;
        const xpath = createTextMatchXpath(action.text);
        console.log(
          `  - Kiểm tra số lượng phần tử chứa text "${action.text}" là ${action.expectedCount}`
        );
        await browser.elements("xpath", xpath, function (res) {
          this.assert.equal(res.value.length, action.expectedCount);
        });
        break;
      }

      case "check_visible": {
        actionTargetText = action.text;
        const xpath = createTextMatchXpathContain(action.text);
        console.log(
          `  - Kiểm tra phần tử chứa text "${action.text}" hiển thị trên giao diện`
        );
        await browser.waitForElementPresent("xpath", xpath, 3000);
        break;
      }

      default:
        throw new Error(`Không hỗ trợ action: ${type}`);
    }

    console.log(`✅ Hoàn thành action: ${type}\n`);
  } catch (error) {
    console.log(`❌ Lỗi khi thực hiện action: ${type} - ${error.message}`);
    throw new Error(`Không thể thực hiện action "${type}": ${error.message}`);
  }
}

module.exports = {
  switchAndRunAction,
};
