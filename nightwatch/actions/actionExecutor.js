const { createTextMatchXpath, createTextMatchXpathContain } = require('../utils/textUtils');

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
  switchAndRunAction,
};
