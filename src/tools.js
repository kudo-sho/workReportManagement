// スプレッドシート初期化
function setupInitialLayout_v2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // まずすべてクリア

  // ヘッダー行を作成
  sheet.appendRow(['日付', '稼働時間（h）', '業務内容', 'イベントID（管理用）']);

  // ヘッダーにスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9ead3'); // 明るい緑

  // 日付列（A列）を日付フォーマットに
  sheet.getRange(2, 1, sheet.getMaxRows() - 1).setNumberFormat('yyyy/MM/dd');

  // 稼働時間列（B列）を小数1桁に
  sheet.getRange(2, 2, sheet.getMaxRows() - 1).setNumberFormat('0.0');

  // 業務内容列（C列）はそのまま

  // イベントID列（D列）を「非表示」にする
  sheet.hideColumn(sheet.getRange(1, 4));
}



