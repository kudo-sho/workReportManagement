// スプレッドシートの設定
const SHEET_NAME = '稼働承認';

// スプレッドシートを取得
function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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

// ユーザー情報を取得
function getUserInfo() {
  const user = Session.getActiveUser();
  return {
    email: user.getEmail(),
    name: user.getUsername()
  };
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
    const approvals = data.slice(1).map(row => {
      const approval = {};
      headers.forEach((header, index) => {
        let value = row[index];
        if (header === 'タイムスタンプ' && value instanceof Date) {
          value = Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
        }
        if (header === '対象月' && value) {
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
    return { headers, approvals };
  } catch (error) {
    console.error('Error fetching work approvals:', error);
    return { headers: [], approvals: [] };
  }
}

// 未承認の月を取得
function getUnapprovedMonths() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName('月次稼働集計表');
    
    if (!summarySheet) {
      console.log('月次稼働集計表が見つかりません');
      return [];
    }

    const lastRow = summarySheet.getLastRow();
    console.log('最終行:', lastRow);
    if (lastRow <= 1) {
      console.log('データが存在しません');
      return [];
    }

    // 月とステータスの列を取得
    const data = summarySheet.getRange(2, 1, lastRow - 1, 6).getValues();
    console.log('取得したデータ:', data);
    
    // 未承認の月をフィルタリング
    const unapprovedMonths = data
      .filter(row => {
        console.log('ステータス確認:', row[5]);
        return row[5] !== '承認済';
      })
      .map(row => {
        let month = row[0];
        // Date型ならYYYY/MMに変換
        if (month instanceof Date) {
          month = Utilities.formatDate(month, 'Asia/Tokyo', 'yyyy/MM');
        }
        console.log('処理中の月:', month);
        const [year, monthNum] = month.split('/');
        return {
          value: `${year}-${monthNum}`,
          label: `${year}年${monthNum}月`
        };
      });

    console.log('未承認の月:', unapprovedMonths);
    return unapprovedMonths;
  } catch (error) {
    console.error('Error fetching unapproved months:', error);
    return [];
  }
} 