/**
 * 環境変数の設定を取得する関数
 * @return {Object|null} 設定オブジェクトまたはnull
 */
function getConfigFromProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const calendarId = scriptProperties.getProperty('CALENDAR_ID');
  const workingStartDate = scriptProperties.getProperty('WORKING_START_DATE');
  
  const missingConfigs = [];
  if (!calendarId) missingConfigs.push('カレンダーID');
  if (!workingStartDate) missingConfigs.push('稼働開始日');
  
  if (missingConfigs.length > 0) {
    const errorMessage = `以下の設定が行われていません：\n${missingConfigs.join('\n')}\n\n「環境設定」から設定を行ってください。`;
    console.error(errorMessage);
    return null;
  }
  
  return {
    calendarId: calendarId,
    workingStartDate: new Date(workingStartDate)
  };
} 