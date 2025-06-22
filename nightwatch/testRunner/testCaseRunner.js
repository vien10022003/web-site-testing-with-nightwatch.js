const { switchAndRunAction } = require('../actions/actionExecutor');

async function runTestCase(actions, expectedList, browser) {
  let i = 0;
  while (i < actions.length) {
    const actionJson = actions[i];
    const actionType = actionJson.action;
    
    // Skip nếu là từ "hoặc"
    if (actionType == "or") {
      console.log(`⚠️ Bỏ qua hành động "${actions[i+1].matchedText}" vì hành động trước đó đã thành công`);
      i = i+2; // Bỏ qua từ "hoặc" và action tiếp theo, vì action trước đó đã thành công
      continue;
    }
    
    global.actionSrcText = actionJson.matchedText;
    
    try {
      await switchAndRunAction(actionJson, actionType, browser);
    } catch (error) {
      // Nếu action fail và có "hoặc" phía sau, thử action tiếp theo
      if (i + 1 < actions.length && actions[i + 1].action === "or") {
        console.log(`⚠️ Action "${actionType}" thất bại, thử alternative sau "hoặc"`);
        // Bỏ qua từ "hoặc" và thử action tiếp theo
        i += 2;
        continue;
      } else {
        // Không có alternative, throw error
        throw error;
      }
    }
    
    i++;
  }

  i = 0;
  // kết quả mong đợi
  while (i < expectedList.length) {
    const actionJson = expectedList[i];
    const actionType = actionJson.action;
    
    // Skip nếu là từ "hoặc"
    if (actionType == "or") {
      console.log(`⚠️ Bỏ qua hành động "${expectedList[i+1].matchedText}" vì hành động trước đó đã thành công`);
      i = i+2; // Bỏ qua từ "hoặc" và action tiếp theo, vì action trước đó đã thành công
      continue;
    }
    
    global.actionSrcText = actionJson.matchedText;
    
    try {
      await switchAndRunAction(actionJson, actionType, browser);
    } catch (error) {
      // Nếu action fail và có "hoặc" phía sau, thử action tiếp theo
      if (i + 1 < expectedList.length && expectedList[i + 1].action === "or") {
        console.log(`⚠️ Action "${actionType}" thất bại, thử alternative sau "hoặc"`);
        // Bỏ qua từ "hoặc" và thử action tiếp theo
        i += 2;
        continue;
      } else {
        // Không có alternative, throw error
        throw error;
      }
    }
    
    i++;
  }

}

module.exports = {
  runTestCase,
};
