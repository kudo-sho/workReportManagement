// スプレッドシートの設定
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_NAME = '稼働承認';

// スプレッドシートを取得
function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    // ヘッダー行を設定
    sheet.appendRow([
      'タイムスタンプ',
      'メールアドレス',
      '氏名',
      '対象月',
      '承認可否',
      'コメント'
    ]);
  }
  
  return sheet;
}

// Webアプリケーションのエントリーポイント
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('稼働承認フォーム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTMLファイルを読み込む
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 稼働承認を送信
function submitWorkApproval(formData) {
  try {
    const sheet = getSheet();
    sheet.appendRow([
      new Date(),
      formData.email,
      formData.name,
      formData.targetMonth,
      formData.approvalStatus,
      formData.comment
    ]);
    return { success: true };
  } catch (error) {
    console.error('Error submitting work approval:', error);
    return { success: false, error: error.toString() };
  }
}

// 稼働承認一覧を取得
function getWorkApprovals() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    // データを新しい順に並べ替え
    const approvals = data.slice(1).reverse().map(row => {
      const approval = {};
      headers.forEach((header, index) => {
        let value = row[index];
        // タイムスタンプのフォーマット
        if (header === 'タイムスタンプ' && value instanceof Date) {
          value = Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
        }
        // 対象月のフォーマット
        if (header === '対象月' && value) {
          // 日付型の場合は年月のみ抽出
          if (value instanceof Date) {
            value = Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM');
          } else if (/^\d{4}-\d{2}/.test(value)) {
            value = value.substring(0, 7);
          }
        }
        approval[header] = value;
      });
      return approval;
    });
    return approvals;
  } catch (error) {
    console.error('Error fetching work approvals:', error);
    return [];
  }
} 