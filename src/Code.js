// スプレッドシートの設定
const SHEET_NAME = '稼働承認一覧';

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

// メール送信関数
function sendWorkApprovalEmail(formData, reportResult = null) {
  const mailBody = `【稼働承認フォーム送信内容】\n\nメールアドレス: ${formData.email}\n氏名: ${formData.name}\n対象月: ${formData.targetMonth}\n承認可否: ${formData.approvalStatus}\nコメント: ${formData.comment || '(なし)'}\n送信日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')}`;
  
  const mailOptions = {
    to: formData.email,
    subject: '稼働承認フォーム送信内容のご案内',
    body: mailBody
  };

  // 月次報告書が作成されていれば添付
  if (reportResult && reportResult.success && reportResult.fileId) {
    try {
      const file = DriveApp.getFileById(reportResult.fileId);
      mailOptions.attachments = [file.getBlob()];
      mailOptions.htmlBody = `${mailBody.replace(/\n/g, '<br>')}<br><br>月次報告書（${reportResult.fileName}）を添付しました。`;
    } catch (e) {
      console.error('メール添付ファイルの取得に失敗しました: ', e);
      // 添付に失敗してもメールは送信する
    }
  }
  
  MailApp.sendEmail(mailOptions);
}

// 稼働承認を送信 (即時処理部分)
function submitWorkApproval(formData) {
  try {
    const sheet = getSheet(); // '稼働承認一覧' シート
    sheet.appendRow([
      new Date(),
      formData.email,
      formData.name,
      formData.targetMonth,
      formData.approvalStatus,
      formData.comment
    ]);

    // 月次稼働集計表のステータスを更新
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName('月次稼働集計表');
    if (summarySheet) {
      const lastRow = summarySheet.getLastRow();
      if (lastRow > 1) {
        const data = summarySheet.getRange(2, 1, lastRow - 1, 1).getValues(); // 月列のみ取得
        for (let i = 0; i < data.length; i++) {
          let month = data[i][0];
          let ym = '';
          if (month instanceof Date) {
            ym = Utilities.formatDate(month, 'Asia/Tokyo', 'yyyy-MM');
          } else if (typeof month === 'string' && month.match(/^\d{4}[\/\-]\d{2}$/)) {
            ym = month.replace('/', '-');
          }
          if (ym === formData.targetMonth) {
            let statusToSet = '';
            if (formData.approvalStatus === '承認') {
              statusToSet = '承認済';
            } else if (formData.approvalStatus === '否認') {
              statusToSet = '否認';
            }
            if (statusToSet) {
              summarySheet.getRange(i + 2, 6).setValue(statusToSet); // 6列目がステータス
            }
            break;
          }
        }
      }
    }
    // メール送信と報告書作成はここでは行わない
    return { success: true, needsReportGeneration: formData.approvalStatus === '承認', formData: formData };
  } catch (error) {
    console.error('Error submitting work approval:', error);
    return { success: false, error: error.toString() };
  }
}

// 月次報告書作成とメール送信 (非同期で呼び出される部分)
function generateReportAndSendEmailAsync(formData) {
  let reportResult = null;
  try {
    if (formData.approvalStatus === '承認') {
      reportResult = makeMonthlyReport(formData.targetMonth); // 報告書作成
    }

    // メール送信 (報告書作成結果を渡す)
    sendWorkApprovalEmail(formData, reportResult);

    return { success: true, reportResult: reportResult, message: "月次報告書の作成とメール送信が完了しました。" };

  } catch (error) {
    console.error('Error in generateReportAndSendEmailAsync:', error);
    // エラー時でも、基本的な承認情報はメールで送る試み (reportResult は null のまま)
    try {
      sendWorkApprovalEmail(formData, null); 
    } catch (emailError) {
      console.error('Error sending basic email after report generation failure:', emailError);
    }
    return { success: false, error: error.toString(), message: "月次報告書の作成またはメール送信中にエラーが発生しました。" };
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

    // タイムスタンプで降順にソート
    approvals.sort((a, b) => {
      const dateA = new Date(a['タイムスタンプ']);
      const dateB = new Date(b['タイムスタンプ']);
      return dateB - dateA; // 降順
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
    
    // 当月（今日の年月）を取得
    const today = new Date();
    const thisMonth = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM');

    // 未承認の月をフィルタリング（当月も除外）
    const unapprovedMonths = data
      .filter(row => {
        console.log('ステータス確認:', row[5]);
        let month = row[0];
        let ym = '';
        if (month instanceof Date) {
          ym = Utilities.formatDate(month, 'Asia/Tokyo', 'yyyy-MM');
        } else if (typeof month === 'string') {
          ym = month.replace('/', '-');
        }
        // ステータスが承認済でなく、かつ当月より前の月のみ
        return row[5] !== '承認済' && ym < thisMonth;
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

// 指定月の稼働内容一覧を取得
function getWorkDetailsByMonth(month) {
  if (!month) return []; // nullや空文字なら即return
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('稼働一覧表');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 検索月を正規化（YYYY-MM形式に統一）
  const normalizedMonth = month.replace('/', '-');
  
  const filtered = data.slice(1).filter(row => {
    let date = row[0];
    if (!date) return false;
    
    let ym = '';
    if (date instanceof Date) {
      ym = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');
    } else if (typeof date === 'string') {
      // 日付文字列を正規化
      const normalized = date.replace(/[.\-]/g, '/');
      const parts = normalized.split('/');
      if (parts.length >= 2) {
        const year = parts[0].replace(/[^0-9]/g, '');
        const monthNum = parts[1].replace(/[^0-9]/g, '').padStart(2, '0');
        ym = `${year}-${monthNum}`;
      }
    }
    
    Logger.log('行の日付: %s, 変換後: %s, 検索月: %s, 一致: %s', 
      date, ym, normalizedMonth, ym === normalizedMonth);
    
    return ym === normalizedMonth;
  });
  
  Logger.log('該当件数: %d', filtered.length);
  
  // データの変換処理を改善
  const result = filtered.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      let value = row[index];
      // 日付型の場合は文字列に変換
      if (value instanceof Date) {
        value = Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd');
      }
      obj[header] = value;
    });
    return obj;
  });
  
  Logger.log('返却データ:', JSON.stringify(result, null, 2));
  return result;
}

// 指定月の月次稼働集計情報を取得
function getMonthlySummaryByMonth(month) {
  if (!month) return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('月次稼働集計表');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 検索月を正規化（YYYY-MM形式に統一）
  const normalizedMonth = month.replace('/', '-');
  
  // 該当月の行を検索
  const targetRow = data.slice(1).find(row => {
    let month = row[0];
    let ym = '';
    if (month instanceof Date) {
      ym = Utilities.formatDate(month, 'Asia/Tokyo', 'yyyy-MM');
    } else if (typeof month === 'string') {
      ym = month.replace('/', '-');
    }
    return ym === normalizedMonth;
  });
  
  if (!targetRow) return null;
  
  return {
    '完了タスク': targetRow[2] || '',
    '未完了及び進行中のタスク': targetRow[3] || '',
    '備考': targetRow[4] || ''
  };
}
