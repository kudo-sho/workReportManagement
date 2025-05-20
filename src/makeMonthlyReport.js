/**
 * 月次作業報告書の生成
 * スプレッドシートのデータを取得して、テンプレートのDocに挿入し、
 * 新しい作業報告書を生成します。
 * 
 * @author 作成日: 2024-05
 * @param {string} targetMonth 対象月 (例: '2024-05')
 */

// ファイルIDを直接指定（ここにテンプレートのファイルIDを設定してください）
// Google Driveでテンプレートファイルを開き、URLの以下の部分からIDを取得：
// https://docs.google.com/document/d/【ここがファイルID】/edit
const TEMPLATE_FILE_ID = '1VxAp2yeCrI06lhvywg06tGUonlPsCkzWSs0J8Nw6g6I'; // ここに実際のIDを入力してください

function makeMonthlyReport(targetMonth = '') {
  try {
    // 対象月が指定されない場合は前月を取得
    if (!targetMonth) {
      const today = new Date();
      today.setMonth(today.getMonth() - 1);
      targetMonth = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM');
    }

    // 年月表示形式を作成（例: 2024-05 → 2024年5月）
    const [year, month] = targetMonth.split('-');
    const displayMonth = `${year}年${parseInt(month)}月`;
    const outputFileName = `${displayMonth}次作業報告書`;

    // スプレッドシートからデータ取得
    const data = getMonthlySummaryData(targetMonth);
    if (!data) {
      throw new Error(`${displayMonth}のデータが見つかりません`);
    }

    // テンプレートファイルを取得 (ファイルIDを直接使用)
    let templateFile = null;
    
    try {
      // ファイルIDが指定されている場合はそれを使用
      if (TEMPLATE_FILE_ID) {
        console.log('テンプレートファイルIDを使用して取得します: ' + TEMPLATE_FILE_ID);
        templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
      } else {
        // IDが指定されていない場合は、名前で検索を試みる
        console.log('テンプレートファイルIDが未設定のため、名前で検索します');
        const docTemplateName = '月次作業報告書(テンプレート)';
        const docTemplateFiles = DriveApp.getFilesByName(docTemplateName);
        
        if (docTemplateFiles && docTemplateFiles.hasNext()) {
          templateFile = docTemplateFiles.next();
          console.log('テンプレートファイルを名前で見つけました: ' + templateFile.getName());
        } else {
          // Word形式で検索
          const wordTemplateName = '月次作業報告書(テンプレート).docx';
          const wordTemplateFiles = DriveApp.getFilesByName(wordTemplateName);
          
          if (wordTemplateFiles && wordTemplateFiles.hasNext()) {
            templateFile = wordTemplateFiles.next();
            console.log('Wordテンプレートファイルを見つけました: ' + templateFile.getName());
          }
        }
      }
    } catch (e) {
      console.error('テンプレートファイル取得エラー: ' + e.message);
    }
    
    // テンプレートが見つからない場合は新規作成
    if (!templateFile) {
      try {
        console.log('テンプレートファイルが見つからないため、新規作成します');
        templateFile = createDefaultTemplate();
        
        if (!templateFile) {
          throw new Error('テンプレートの作成に失敗しました');
        }
      } catch (e) {
        throw new Error('テンプレートファイルが見つからず、作成もできませんでした: ' + e.message);
      }
    }
    
    // テンプレートをGoogle Docsとしてコピー
    let newDoc = null;
    if (templateFile.getMimeType() === 'application/vnd.google-apps.document') {
      // Google Doc形式のテンプレートならそのままコピー
      console.log('Google Doc形式のテンプレートを使用します');
      const newDocFile = templateFile.makeCopy(outputFileName);
      console.log('テンプレートをコピーしました。ID: ' + newDocFile.getId());
      newDoc = DocumentApp.openById(newDocFile.getId());
    } else {
      // WordファイルならGoogle Doc形式に変換
      console.log('Word形式のテンプレートを使用します: ' + templateFile.getName() + ', MIME: ' + templateFile.getMimeType());
      newDoc = DocumentApp.create(outputFileName);
      console.log('空のGoogle Docを作成しました。ID: ' + newDoc.getId());
      
      // テンプレートの内容を確認
      try {
        const templateContent = templateFile.getBlob().getDataAsString();
        console.log('テンプレートの内容（最初の100文字）: ' + templateContent.substring(0, 100) + '...');
      } catch (te) {
        console.log('テンプレート内容の取得に失敗: ' + te.message);
      }
      // 注意: ここではWordファイルの内容を直接反映できない
      // ユーザーにはGoogle Doc形式のテンプレートを用意してもらう
    }
    
    // テキスト置換
    const body = newDoc.getBody();
    console.log('テキスト置換前のドキュメント内容（最初の100文字）: ' + body.getText().substring(0, 100) + '...');
    
    // プレースホルダーの存在を確認
    const text = body.getText();
    for (const key in data) {
      if (text.indexOf(`{{${key}}}`) !== -1) {
        console.log(`プレースホルダー {{${key}}} が見つかりました`);
      } else {
        console.log(`プレースホルダー {{${key}}} が見つかりません`);
      }
    }
    
    // テキスト置換を実行
    for (const key in data) {
      if (typeof data[key] === 'string' || typeof data[key] === 'number') {
        const replacementValue = data[key].toString();
        console.log(`置換: {{${key}}} → ${replacementValue.substring(0, 30)}${replacementValue.length > 30 ? '...' : ''}`);
        body.replaceText(`{{${key}}}`, replacementValue);
      }
    }
    
    console.log('テキスト置換後のドキュメント内容（最初の100文字）: ' + body.getText().substring(0, 100) + '...');
    newDoc.saveAndClose();
    
    // Google DocのIDを取得
    const docId = newDoc.getId();
    const docFile = DriveApp.getFileById(docId);
    
    console.log('Google Docファイルを作成しました。ID: ' + docId);
    console.log('コンテンツの確認: ' + DocumentApp.openById(docId).getBody().getText().substring(0, 100) + '...');
    
    // Method 1: Drive APIのAdvancedサービスを使用してエクスポート
    try {
      // Drive v2 APIを使用
      if (typeof Drive !== 'undefined') {
        console.log('Method 1: Drive APIを使用してエクスポートを試みます...');
        const wordBlob = Drive.Files.export(docId, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', {alt: 'media'});
        console.log('エクスポート成功。BLOBサイズ: ' + (wordBlob ? wordBlob.getBytes().length : 'null'));
        
        if (!wordBlob || wordBlob.getBytes().length === 0) {
          throw new Error('エクスポートされたファイルが空です');
        }
        
        const wordFile = DriveApp.createFile(wordBlob);
        wordFile.setName(`${outputFileName}.docx`);
        
        console.log('Wordファイルを作成しました。名前: ' + wordFile.getName() + ', サイズ: ' + wordFile.getSize() + ' bytes');
        
        // 生成したファイルのURLを返す
        return { 
          success: true, 
          message: `${displayMonth}の月次作業報告書を生成しました`, 
          fileUrl: wordFile.getUrl(),
          fileName: wordFile.getName()
        };
      } else {
        console.log('Drive APIが利用できません。Method 2に進みます...');
        throw new Error('Drive APIが利用できません');
      }
    } catch (e) {
      console.error('Word変換エラー (Method 1): ' + e.message);
      
      // Method 2: URLを使用してエクスポート
      try {
        console.log('Method 2: URL Fetchを使用してエクスポートを試みます...');
        const token = ScriptApp.getOAuthToken();
        const url = `https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document`;
        
        const response = UrlFetchApp.fetch(url, {
          headers: {
            'Authorization': 'Bearer ' + token
          },
          muteHttpExceptions: true // エラーを詳細に捕捉するため
        });
        
        const responseCode = response.getResponseCode();
        console.log('URL Fetchレスポンスコード: ' + responseCode);
        
        if (responseCode !== 200) {
          throw new Error('エクスポート失敗: レスポンスコード ' + responseCode);
        }
        
        const blob = response.getBlob();
        console.log('取得したBLOBサイズ: ' + (blob ? blob.getBytes().length : 'null'));
        
        if (!blob || blob.getBytes().length === 0) {
          throw new Error('エクスポートされたファイルが空です');
        }
        
        blob.setName(`${outputFileName}.docx`);
        const wordFile = DriveApp.createFile(blob);
        
        console.log('Wordファイルを作成しました。名前: ' + wordFile.getName() + ', サイズ: ' + wordFile.getSize() + ' bytes');
        
        // 生成したファイルのURLを返す
        return { 
          success: true, 
          message: `${displayMonth}の月次作業報告書を生成しました`, 
          fileUrl: wordFile.getUrl(),
          fileName: wordFile.getName()
        };
      } catch (e2) {
        console.error('Word変換エラー (Method 2): ' + e2.message);
        
        // Method 3: ダウンロードリンクを提供
        console.log('Method 3: ダウンロードリンクを提供します...');
        const exportUrl = `https://docs.google.com/document/d/${docId}/export?format=docx`;
        
        return {
          success: true,
          message: `${displayMonth}の月次作業報告書を生成しました（Word形式での直接ダウンロードリンク）`, 
          fileUrl: exportUrl,
          fileName: `${outputFileName}.docx`
        };
      }
    }
    
  } catch (error) {
    console.error(`月次作業報告書生成エラー: ${error.message}`);
    return { success: false, message: error.message };
  }
}

/**
 * デフォルトのテンプレートを作成する
 * @return {DriveApp.File} 作成したテンプレートファイル
 */
function createDefaultTemplate() {
  try {
    // デフォルトのテンプレートを作成
    const templateName = '月次作業報告書(テンプレート)';
    const doc = DocumentApp.create(templateName);
    const body = doc.getBody();
    
    // ヘッダー追加
    const header = doc.addHeader();
    header.appendParagraph('月次作業報告書').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    // タイトル
    body.appendParagraph('{{month}} 作業報告書').setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // プロジェクト情報テーブル
    const projectTable = body.appendTable([
      ['プロジェクト名', '{{projectName}}'],
      ['クライアント名', '{{clientName}}'],
      ['担当者名', '{{employeeName}}'],
      ['管理者名', '{{managerName}}']
    ]);
    projectTable.setAttributes({
      'borderWidth': 1,
      'width': 500
    });
    
    // 稼働情報テーブル
    body.appendParagraph('稼働情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const workTable = body.appendTable([
      ['稼働日数', '{{totalWorkDays}}日'],
      ['実稼働日数', '{{actualWorkDays}}日'],
      ['稼働時間', '{{workingHours}}時間']
    ]);
    workTable.setAttributes({
      'borderWidth': 1,
      'width': 500
    });
    
    // 業務概要
    body.appendParagraph('業務概要').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('{{workSummary}}');
    
    // 詳細
    body.appendParagraph('作業詳細').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('{{workDetails}}');
    
    // フッター追加
    const footer = doc.addFooter();
    footer.appendParagraph('作成日: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
    
    doc.saveAndClose();
    
    console.log('デフォルトテンプレートを作成しました。ID: ' + doc.getId());
    // テンプレートファイルIDをログに出力（次回ファイルIDを設定するため）
    console.log('次回以降は以下のファイルIDを定数に設定してください: ' + doc.getId());
    
    return DriveApp.getFileById(doc.getId());
  } catch (e) {
    console.error('デフォルトテンプレート作成エラー: ' + e.message);
    return null;
  }
}

/**
 * 指定月の稼働データをスプレッドシートから取得
 * @param {string} targetMonth 対象月 (例: '2024-05')
 * @return {Object} 置換用データオブジェクト
 */
function getMonthlySummaryData(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('月次稼働集計表');
  
  if (!summarySheet) {
    throw new Error('月次稼働集計表が見つかりません');
  }
  
  const lastRow = summarySheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('月次稼働集計表にデータがありません');
  }
  
  const data = summarySheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  // 指定月のデータを検索
  let monthData = null;
  for (let i = 0; i < data.length; i++) {
    let month = data[i][0]; // 月列
    let monthStr = '';
    
    if (month instanceof Date) {
      monthStr = Utilities.formatDate(month, 'Asia/Tokyo', 'yyyy-MM');
    } else if (typeof month === 'string') {
      monthStr = month.replace('/', '-');
    }
    
    if (monthStr === targetMonth) {
      monthData = {
        targetMonth: Utilities.formatDate(new Date(`${targetMonth}-01`), 'Asia/Tokyo', 'yyyy年M月'),
        totalWorkDays: data[i][1] || 0,
        actualWorkDays: data[i][2] || 0,
        workingHours: data[i][3] || 0,
        workSummary: data[i][4] || '',
        status: data[i][5] || '',
        projectName: data[i][6] || '',
        clientName: data[i][7] || '',
        employeeName: data[i][8] || '',
        managerName: data[i][9] || ''
      };
      break;
    }
  }
  
  // 詳細データを取得
  if (monthData) {
    // 作業内容詳細を取得
    const details = getWorkDetailsByMonth(targetMonth);
    monthData.workDetails = details.map((detail, i) => {
      return `${i+1}. ${detail.workContent || ''}`;
    }).join('\n');
  }
  
  return monthData;
}

/**
 * UI側から呼び出すための関数
 */
function generateMonthlyReport(month) {
  return makeMonthlyReport(month);
}

/**
 * スプレッドシートのカスタムメニューに追加するための関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('月次報告')
    .addItem('月次作業報告書を生成', 'showReportGeneratorDialog')
    .addToUi();
}

/**
 * 報告書生成ダイアログを表示
 */
function showReportGeneratorDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        margin: 20px;
      }
      .form-group {
        margin-bottom: 15px;
      }
      label {
        display: block;
        margin-bottom: 5px;
        font-weight: bold;
      }
      select {
        width: 100%;
        padding: 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
      }
      button {
        background-color: #4285f4;
        color: white;
        border: none;
        padding: 10px 15px;
        border-radius: 4px;
        cursor: pointer;
      }
      .result {
        margin-top: 20px;
        padding: 10px;
        border: 1px solid #ddd;
        border-radius: 4px;
        display: none;
      }
      .success {
        background-color: #d4edda;
      }
      .error {
        background-color: #f8d7da;
      }
    </style>
    
    <div class="form-group">
      <label for="month">対象月を選択</label>
      <select id="month"></select>
    </div>
    
    <button onclick="generateReport()">報告書を生成</button>
    
    <div id="result" class="result"></div>
    
    <script>
      // 初期化
      google.script.run
        .withSuccessHandler(loadMonths)
        .withFailureHandler(showError)
        .getMonthsForReport();
      
      function loadMonths(months) {
        const select = document.getElementById('month');
        
        if (months && months.length) {
          months.forEach(function(item) {
            const option = document.createElement('option');
            option.value = item.value;
            option.textContent = item.label;
            select.appendChild(option);
          });
        } else {
          const option = document.createElement('option');
          option.textContent = '対象月がありません';
          select.appendChild(option);
          select.disabled = true;
          document.querySelector('button').disabled = true;
        }
      }
      
      function generateReport() {
        const month = document.getElementById('month').value;
        const resultDiv = document.getElementById('result');
        
        resultDiv.innerHTML = '処理中...';
        resultDiv.className = 'result';
        resultDiv.style.display = 'block';
        
        google.script.run
          .withSuccessHandler(function(result) {
            if (result.success) {
              resultDiv.className = 'result success';
              resultDiv.innerHTML = result.message + '<br><br>' + 
                '<a href="' + result.fileUrl + '" target="_blank">' + 
                result.fileName + 'を開く</a>';
            } else {
              showError(result.message);
            }
          })
          .withFailureHandler(showError)
          .generateMonthlyReport(month);
      }
      
      function showError(error) {
        const resultDiv = document.getElementById('result');
        resultDiv.className = 'result error';
        resultDiv.innerHTML = 'エラー: ' + error;
        resultDiv.style.display = 'block';
      }
    </script>
  `)
  .setWidth(400)
  .setHeight(300)
  .setTitle('月次作業報告書ジェネレーター');
  
  SpreadsheetApp.getUi().showModalDialog(html, '月次作業報告書ジェネレーター');
}

/**
 * 報告書生成用の月リストを取得
 */
function getMonthsForReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('月次稼働集計表');
    
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const monthData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const months = [];
    
    monthData.forEach(function(row) {
      let month = row[0];
      let monthStr = '';
      let displayMonth = '';
      
      if (month instanceof Date) {
        monthStr = Utilities.formatDate(month, 'Asia/Tokyo', 'yyyy-MM');
        displayMonth = Utilities.formatDate(month, 'Asia/Tokyo', 'yyyy年M月');
      } else if (typeof month === 'string') {
        monthStr = month.replace('/', '-');
        const [y, m] = month.split(/[\/\-]/);
        displayMonth = `${y}年${parseInt(m)}月`;
      }
      
      if (monthStr) {
        months.push({
          value: monthStr,
          label: displayMonth
        });
      }
    });
    
    return months;
  } catch (error) {
    console.error('月リスト取得エラー:', error);
    return [];
  }
}
