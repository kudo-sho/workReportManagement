// テスト関数
function testGetWorkDetailsByMonth() {
  // テストケース1: 正常系 - Date型の日付
  const testDate1 = new Date('2025-04-15');
  const testMonth1 = '2025-04';
  Logger.log('テストケース1: Date型の日付');
  Logger.log('期待値: 2025-04のデータが取得できること');
  const result1 = getWorkDetailsByMonth(testMonth1);
  Logger.log('結果:', result1);

  // テストケース2: 正常系 - 文字列型の日付（YYYY/MM/DD形式）
  const testDate2 = '2025/04/20';
  const testMonth2 = '2025-04';
  Logger.log('テストケース2: 文字列型の日付（YYYY/MM/DD形式）');
  Logger.log('期待値: 2025-04のデータが取得できること');
  const result2 = getWorkDetailsByMonth(testMonth2);
  Logger.log('結果:', result2);

  // テストケース3: 正常系 - 文字列型の日付（YYYY-MM-DD形式）
  const testDate3 = '2025-04-25';
  const testMonth3 = '2025-04';
  Logger.log('テストケース3: 文字列型の日付（YYYY-MM-DD形式）');
  Logger.log('期待値: 2025-04のデータが取得できること');
  const result3 = getWorkDetailsByMonth(testMonth3);
  Logger.log('結果:', result3);

  // テストケース4: 異常系 - 存在しない月
  const testMonth4 = '2025-13';
  Logger.log('テストケース4: 存在しない月');
  Logger.log('期待値: 空配列が返されること');
  const result4 = getWorkDetailsByMonth(testMonth4);
  Logger.log('結果:', result4);

  // テストケース5: 異常系 - 不正な形式の月
  const testMonth5 = '2025/04';
  Logger.log('テストケース5: 不正な形式の月');
  Logger.log('期待値: 空配列が返されること');
  const result5 = getWorkDetailsByMonth(testMonth5);
  Logger.log('結果:', result5);

  // テストケース6: 異常系 - null
  const testMonth6 = null;
  Logger.log('テストケース6: null');
  Logger.log('期待値: 空配列が返されること');
  const result6 = getWorkDetailsByMonth(testMonth6);
  Logger.log('結果:', result6);
}

// テストデータのセットアップ
function setupTestData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('稼働一覧表');
  if (!sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = ss.insertSheet('稼働一覧表');
    // ヘッダー行を設定
    sheet.appendRow([
      '日付',
      'プロジェクト名',
      '作業内容',
      '工数'
    ]);
  }

  // テストデータを追加
  const testData = [
    [new Date('2025-04-15'), 'プロジェクトA', '要件定義', 8],
    ['2025/04/20', 'プロジェクトB', '設計', 6],
    ['2025-04-25', 'プロジェクトC', '実装', 4],
    [new Date('2025-05-01'), 'プロジェクトD', 'テスト', 8]
  ];

  testData.forEach(row => {
    sheet.appendRow(row);
  });
}

// テストの実行
function runTests() {
  setupTestData();
  testGetWorkDetailsByMonth();
} 