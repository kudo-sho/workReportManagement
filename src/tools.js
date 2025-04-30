// スプレッドシート初期化
function setupInitialLayout() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('稼働一覧表');
  if (!sheet) {
    console.error('「稼働一覧表」シートが見つかりません。');
    return;
  }
  sheet.clear(); // まずすべてクリア

  // ヘッダー行を作成
  sheet.appendRow(['日付', '稼働時間（h）', '業務内容', 'ステータス', 'イベントID（管理用）']);

  // 空の表を作成（20行分）
  const emptyRows = Array(20).fill(['', '', '', '', '']);
  sheet.getRange(2, 1, 20, 5).setValues(emptyRows);

  // 列の配置を設定
  sheet.getRange('A:B').setHorizontalAlignment('center'); // 日付と稼働時間は中央揃え
  sheet.getRange('C:C').setHorizontalAlignment('left');   // 業務内容は左揃え
  sheet.getRange('D:E').setHorizontalAlignment('center'); // ステータスとイベントIDは中央揃え
  sheet.getRange('A:E').setVerticalAlignment('middle');   // すべての列を垂直方向中央揃え

  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9ead3'); // 明るい緑
  headerRange.setHorizontalAlignment('center'); // ヘッダーは中央揃え

  // イベントID列（E列）を「非表示」にする
  sheet.hideColumn(sheet.getRange(1, 5));

  // 列幅を設定
  sheet.setColumnWidth(1, 100);  // 日付
  sheet.setColumnWidth(2, 100);  // 稼働時間
  sheet.setColumnWidth(3, 400);  // 業務内容を広めに
  sheet.setColumnWidth(4, 100);  // ステータス

  // ヘッダー行を固定
  sheet.setFrozenRows(1);

  // 業務内容列の文字列を折り返す設定
  sheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}
