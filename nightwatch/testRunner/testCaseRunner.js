const { switchAndRunAction } = require('../actions/actionExecutor');

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

module.exports = {
  runTestCase,
};
