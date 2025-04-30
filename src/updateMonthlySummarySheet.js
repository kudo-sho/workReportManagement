// 月次稼働集計表を作成&更新
function updateMonthlySummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = '稼働一覧表';
  const summarySheetName = '月次稼働集計表';

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`元データシート '${sourceSheetName}' が見つかりません`);
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow <= 1) return; // データなし

  const data = sourceSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A列(日付)・B列(稼働時間)

  // 月単位で合計するためのMapを作る
  const monthlySummary = {};
  data.forEach(row => {
    const date = row[0];
    const hours = row[1];

    if (date instanceof Date && typeof hours === 'number') {
      const ym = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM');
      if (!monthlySummary[ym]) {
        monthlySummary[ym] = 0;
      }
      monthlySummary[ym] += hours;
    }
  });

  // 出力シートが存在しなければ新規作成
  let summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(summarySheetName);
    setupMonthlySummaryLayout();
  }

  // 既存のデータを取得
  const lastDataRow = summarySheet.getLastRow();
  const existingData = lastDataRow > 1 ? summarySheet.getRange(2, 1, lastDataRow - 1, 5).getValues() : [];

  // 月順に並べて出力（昇順）
  const sortedMonths = Object.keys(monthlySummary).sort();
  sortedMonths.forEach((ym, index) => {
    const row = index + 2; // ヘッダー行の次から開始
    const existingRow = existingData.find(row => row[0] === ym);
    
    if (existingRow) {
      // 既存の行がある場合は稼働時間のみ更新
      summarySheet.getRange(row, 2).setValue(Math.round(monthlySummary[ym] * 10) / 10);
    } else {
      // 新しい行の場合は追加
      summarySheet.getRange(row, 1).setValue(ym);
      summarySheet.getRange(row, 2).setValue(Math.round(monthlySummary[ym] * 10) / 10);
    }
  });

  // 書式設定
  summarySheet.getRange('B2:B' + (sortedMonths.length + 1)).setNumberFormat('0.0');
}