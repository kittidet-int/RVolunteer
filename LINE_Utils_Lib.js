function sendMessage(targetId, type, altText, message) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lineAccessToken = scriptProperties.getProperty('LINE_ACCESS_TOKEN');

  if (!lineAccessToken || !targetId) {
    throw new Error("[LINE] Access Token or Target ID not found");
  }

  var payload = {
    'to': targetId,
    'messages': [{
      'type': type,
      'altText': altText,
      'contents': message
    }]
  };

  var options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + lineAccessToken
    },
    'payload': JSON.stringify(payload)
  };

  try {
    var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
    Logger.log('[LINE] Success');
  } catch (e) {
    throw new Error("[LINE] Error: " + e.toString());
  }
}

function sendMessageToGroup(type, altText, message) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lineTargetId = scriptProperties.getProperty('LINE_TARGET_ID');

  sendMessage(lineTargetId, type, altText, message);
}

function sendMessageToTest(type, altText, message) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lineTestId = scriptProperties.getProperty('LINE_TEST_ID');

  sendMessage(lineTestId, type, altText, message);
}
