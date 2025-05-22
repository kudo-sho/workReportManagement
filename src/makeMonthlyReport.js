/**
 * 月次作業報告書の生成
 * スプレッドシートのデータを取得して、テンプレートのDocに挿入し、
 * 新しい作業報告書を生成します。
 * 
 * @author 作成日: 2024-05
 * @param {string} targetMonth 対象月 (例: '2024-05')
 * @param {boolean} exportWord Google DocをWordに変換するか（デフォルト: false）
 * @param {string} outputFolderId 出力先フォルダのID（指定しない場合はルートフォルダに出力）
 */

// ファイルIDを直接指定（ここにテンプレートのファイルIDを設定してください）
// Google Driveでテンプレートファイルを開き、URLの以下の部分からIDを取得：
// https://docs.google.com/document/d/【ここがファイルID】/edit
const TEMPLATE_FILE_ID = '1Xrr-Eg9JlGJZrwWODDEzD6O_UIOhDY9QHQnRl6LK2qI'; // ここに実際のIDを入力してください

function makeMonthlyReport(targetMonth = '', exportWord = false, outputFolderId = '1S6aH6ZmwVYgG8WldgxNpZXs-Zt_C_fY0') {
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
  
    
    // テンプレートをGoogle Docsとしてコピー
    let newDoc = null;
    let newDocFile = null;
    
    if (templateFile.getMimeType() === 'application/vnd.google-apps.document') {
      // Google Doc形式のテンプレートならそのままコピー
      console.log('Google Doc形式のテンプレートを使用します');
      
      // 出力先フォルダが指定されている場合は、そのフォルダにコピー
      if (outputFolderId) {
        try {
          const targetFolder = DriveApp.getFolderById(outputFolderId);
          if (targetFolder) {
            // フォルダに直接コピー
            newDocFile = templateFile.makeCopy(outputFileName, targetFolder);
            console.log('テンプレートを指定フォルダにコピーしました。ID: ' + newDocFile.getId());
          }
        } catch (folderError) {
          console.error('出力先フォルダ取得エラー: ' + folderError.message);
          // フォルダ取得に失敗した場合は、通常のコピーを実施
          newDocFile = templateFile.makeCopy(outputFileName);
          console.log('テンプレートをルートフォルダにコピーしました。ID: ' + newDocFile.getId());
        }
      } else {
        // 出力先フォルダ未指定の場合
        newDocFile = templateFile.makeCopy(outputFileName);
        console.log('テンプレートをルートフォルダにコピーしました。ID: ' + newDocFile.getId());
      }
      
      newDoc = DocumentApp.openById(newDocFile.getId());
    } else {
      // WordファイルならGoogle Doc形式に変換
      console.log('Word形式のテンプレートを使用します: ' + templateFile.getName() + ', MIME: ' + templateFile.getMimeType());
      newDoc = DocumentApp.create(outputFileName);
      newDocFile = DriveApp.getFileById(newDoc.getId());
      console.log('空のGoogle Docを作成しました。ID: ' + newDoc.getId());
      
      // 出力先フォルダが指定されている場合は、そのフォルダに移動
      if (outputFolderId) {
        try {
          const targetFolder = DriveApp.getFolderById(outputFolderId);
          if (targetFolder) {
            // 作成したファイルを指定フォルダに移動
            targetFolder.addFile(newDocFile);
            DriveApp.getRootFolder().removeFile(newDocFile); // ルートフォルダから削除
            console.log('空のGoogle Docを指定フォルダに移動しました');
          }
        } catch (folderError) {
          console.error('出力先フォルダ取得エラー: ' + folderError.message);
          // 処理を続行
        }
      }
      
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
    
    // テーブルの挿入処理（テキスト置換の前に行う）
    if (data.details && data.details.length > 0) {
      insertWorkDetailsTable(body, data.details);
    }
    
    // テキスト置換を実行（テーブル以外の置換）
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
    
    // Google Docのみを返す（Word変換が不要な場合）
    if (!exportWord) {
      return { 
        success: true, 
        message: `${displayMonth}の月次作業報告書を生成しました（Google Doc形式）`, 
        fileUrl: docFile.getUrl(),
        fileName: docFile.getName()
      };
    }
    
    // 以下はWord形式への変換処理（exportWord=trueの場合のみ実行）
    console.log('Word形式への変換を開始します...');
    
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
        
        // Wordファイルを作成（出力先フォルダに保存）
        let wordFile = null;
        if (outputFolderId) {
          try {
            const targetFolder = DriveApp.getFolderById(outputFolderId);
            wordBlob.setName(`${outputFileName}.docx`);
            wordFile = targetFolder.createFile(wordBlob);
            console.log('Wordファイルを指定フォルダに作成しました。名前: ' + wordFile.getName());
          } catch (folderError) {
            console.error('出力先フォルダへのWord保存エラー: ' + folderError.message);
            // フォルダ取得に失敗した場合は、ルートに保存
            wordBlob.setName(`${outputFileName}.docx`);
            wordFile = DriveApp.createFile(wordBlob);
            console.log('Wordファイルをルートフォルダに作成しました。名前: ' + wordFile.getName());
          }
        } else {
          // 出力先フォルダ未指定の場合はルートに保存
          wordBlob.setName(`${outputFileName}.docx`);
          wordFile = DriveApp.createFile(wordBlob);
          console.log('Wordファイルをルートフォルダに作成しました。名前: ' + wordFile.getName());
        }
        
        console.log('Wordファイルを作成しました。名前: ' + wordFile.getName() + ', サイズ: ' + wordFile.getSize() + ' bytes');
        
        // 生成したファイルのURLを返す
        return { 
          success: true, 
          message: `${displayMonth}の月次作業報告書を生成しました（Word形式）`, 
          fileUrl: wordFile.getUrl(),
          fileName: wordFile.getName(),
          googleDocUrl: docFile.getUrl(),
          googleDocName: docFile.getName()
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
        
        // Wordファイルを作成（出力先フォルダに保存）
        let wordFile = null;
        blob.setName(`${outputFileName}.docx`);
        
        if (outputFolderId) {
          try {
            const targetFolder = DriveApp.getFolderById(outputFolderId);
            wordFile = targetFolder.createFile(blob);
            console.log('Wordファイルを指定フォルダに作成しました。名前: ' + wordFile.getName());
          } catch (folderError) {
            console.error('出力先フォルダへのWord保存エラー: ' + folderError.message);
            // フォルダ取得に失敗した場合は、ルートに保存
            wordFile = DriveApp.createFile(blob);
            console.log('Wordファイルをルートフォルダに作成しました。名前: ' + wordFile.getName());
          }
        } else {
          // 出力先フォルダ未指定の場合はルートに保存
          wordFile = DriveApp.createFile(blob);
          console.log('Wordファイルをルートフォルダに作成しました。名前: ' + wordFile.getName());
        }
        
        console.log('Wordファイルを作成しました。名前: ' + wordFile.getName() + ', サイズ: ' + wordFile.getSize() + ' bytes');
        
        // 生成したファイルのURLを返す
        return { 
          success: true, 
          message: `${displayMonth}の月次作業報告書を生成しました（Word形式）`, 
          fileUrl: wordFile.getUrl(),
          fileName: wordFile.getName(),
          googleDocUrl: docFile.getUrl(),
          googleDocName: docFile.getName()
        };
      } catch (e2) {
        console.error('Word変換エラー (Method 2): ' + e2.message);
        
        // Method 3: ダウンロードリンクを提供
        console.log('Method 3: ダウンロードリンクを提供します...');
        const exportUrl = `https://docs.google.com/document/d/${docId}/export?format=docx`;
        
        return {
          success: true,
          message: `${displayMonth}の月次作業報告書を生成しました（Google Doc形式と直接ダウンロードリンク）`, 
          fileUrl: docFile.getUrl(),
          fileName: docFile.getName(),
          wordExportUrl: exportUrl,
          wordFileName: `${outputFileName}.docx`
        };
      }
    }
    
  } catch (error) {
    console.error(`月次作業報告書生成エラー: ${error.message}`);
    return { success: false, message: error.message };
  }
}

/**
 * 指定月の稼働一覧表データをスプレッドシートから取得
 * @param {string} targetMonth 対象月 (例: '2024-05')
 * @return {Array} 稼働一覧表データの配列
 */
function getWorkDetailsByMonthForMakeReport(targetMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailSheet = ss.getSheetByName('稼働一覧表');
    
    if (!detailSheet) {
      console.log('稼働一覧表シートが見つかりません');
      return [];
    }
    
    const lastRow = detailSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('稼働一覧表シートにデータがありません');
      return [];
    }
    
    // データ範囲: 日付, 稼働時間(h), 業務内容, ステータス
    const data = detailSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    
    // 指定月のデータをフィルタリング
    const details = [];
    data.forEach(function(row) {
      let workDate = row[0];
      let workTime = row[1] || 0;
      let workContent = row[2] || '';
      let status = row[3] || '';
      
      // 日付データをフォーマット
      let dateStr = '';
      if (workDate instanceof Date) {
        dateStr = Utilities.formatDate(workDate, 'Asia/Tokyo', 'yyyy-MM-dd');
        
        // 対象月のデータのみ抽出
        if (dateStr.startsWith(targetMonth)) {
          const formattedDate = Utilities.formatDate(workDate, 'Asia/Tokyo', 'MM/dd');
          details.push({
            workDate: formattedDate,
            workTime: workTime,
            workContent: workContent,
            status: status
          });
        }
      }
    });
    
    console.log(`${targetMonth}の稼働一覧表データを${details.length}件取得しました`);
    return details;
    
  } catch (error) {
    console.error('稼働一覧表データ取得エラー:', error);
    return [];
  }
}

/**
 * Googleドキュメントに稼働詳細テーブルを作成する
 * @param {DocumentApp.Body} body ドキュメントのボディ
 * @param {Array} details 稼働詳細データの配列
 */
function insertWorkDetailsTable(body, details) {
  try {
    // テーブルのプレースホルダーを探す
    const text = body.getText();
    
    // まず既存のテーブルプレースホルダーを探して置換
    if (text.indexOf('{{work_details_table}}') !== -1) {
      console.log('稼働詳細テーブルプレースホルダーを見つけました');
      
      // プレースホルダーのある段落を見つける
      const paragraphs = body.getParagraphs();
      let targetParagraph = null;
      let targetIndex = -1;
      
      for (let i = 0; i < paragraphs.length; i++) {
        if (paragraphs[i].getText().indexOf('{{work_details_table}}') !== -1) {
          targetParagraph = paragraphs[i];
          targetIndex = i;
          break;
        }
      }
      
      if (targetParagraph) {
        // テーブルデータを作成
        const tableData = createWorkDetailsTable(details);
        
        if (tableData.length > 1) { // ヘッダーだけでなく少なくとも1行のデータがある
          // プレースホルダーテキストを空白に置き換え（完全に空にはしない）
          targetParagraph.setText(" ");
          
          try {
            // テーブルを挿入
            const table = body.insertTable(targetIndex, tableData);
            console.log('稼働詳細テーブルを挿入しました。行数: ' + table.getNumRows());
            
            // 元の段落を削除（テーブル挿入後）
            body.removeChild(targetParagraph);
            console.log('プレースホルダー段落を削除しました');
          } catch (innerError) {
            console.error('テーブル挿入中のエラー: ' + innerError.message);
            
            // 代替方法: 既存の段落にテキスト形式でデータを挿入
            let textTable = "稼働詳細:\n";
            for (let i = 1; i < tableData.length; i++) { // ヘッダー行はスキップ
              textTable += tableData[i][0] + ": " + tableData[i][2] + " (" + tableData[i][1] + "h)\n";
            }
            targetParagraph.setText(textTable);
          }
        } else {
          // データがない場合
          targetParagraph.setText("稼働詳細データがありません");
        }
      }
    } else {
      // テーブルリテラル形式のプレースホルダー（{table}...{/table}）を探す
      console.log('テーブルプレースホルダーを検索します');
      
      // 正規表現を使用してプレースホルダーを探すことはできないため、
      // テーブルマーカーを探して処理します
      const tableStartMarker = '{{table_start}}';
      const tableEndMarker = '{{table_end}}';
      
      if (text.indexOf(tableStartMarker) !== -1 && text.indexOf(tableEndMarker) !== -1) {
        console.log('テーブルマーカーを見つけました');
        
        // テーブルの位置を特定
        const paragraphs = body.getParagraphs();
        let startIndex = -1;
        let endIndex = -1;
        
        for (let i = 0; i < paragraphs.length; i++) {
          const paragraphText = paragraphs[i].getText();
          if (paragraphText.indexOf(tableStartMarker) !== -1) {
            startIndex = i;
          }
          if (paragraphText.indexOf(tableEndMarker) !== -1) {
            endIndex = i;
            break;
          }
        }
        
        if (startIndex !== -1 && endIndex !== -1) {
          // マーカーを削除（完全に空にはしない）
          paragraphs[startIndex].setText(paragraphs[startIndex].getText().replace(tableStartMarker, " "));
          paragraphs[endIndex].setText(paragraphs[endIndex].getText().replace(tableEndMarker, " "));
          
          // この間に稼働詳細テーブルを挿入
          const tableData = createWorkDetailsTable(details);
          
          if (tableData.length > 1) { // ヘッダーだけでなく少なくとも1行のデータがある
            try {
              const table = body.insertTable(startIndex + 1, tableData);
              console.log('稼働詳細テーブルを挿入しました。行数: ' + table.getNumRows());
            } catch (innerError) {
              console.error('テーブル挿入中のエラー: ' + innerError.message);
              
              // 代替方法: テキスト形式でデータを挿入
              let textTable = "稼働詳細:\n";
              for (let i = 1; i < tableData.length; i++) { // ヘッダー行はスキップ
                textTable += tableData[i][0] + ": " + tableData[i][2] + " (" + tableData[i][1] + "h)\n";
              }
              body.insertParagraph(startIndex + 1, textTable);
            }
          } else {
            // データがない場合
            body.insertParagraph(startIndex + 1, "稼働詳細データがありません");
          }
        }
      } else {
        console.log('テーブルマーカーが見つかりません。テーブルを挿入するには、テンプレートに {{work_details_table}} または {{table_start}} と {{table_end}} マーカーを追加してください。');
      }
    }
  } catch (error) {
    console.error('テーブル挿入エラー: ' + error.message);
  }
}

/**
 * 稼働詳細テーブルを作成する
 * @param {Array} details 稼働詳細データの配列
 * @return {Array} テーブルのセルデータの二次元配列
 */
function createWorkDetailsTable(details) {
  // テーブルのセルデータを作成
  const tableData = [['日付', '稼働時間(hours)', '業務内容']];
  
  // 詳細データを行として追加
  if (details && details.length > 0) {
    details.forEach(function(detail) {
      // nullや未定義の値を処理（空文字列に変換）
      const date = detail.workDate || " ";
      const time = detail.workTime ? detail.workTime.toString() : "0";
      const content = detail.workContent || " ";
      
      tableData.push([date, time, content]);
    });
  } else {
    // データがない場合でも空の行を1つ追加（完全に空のテーブルは作成できないため）
    tableData.push([" ", " ", " "]);
  }
  
  return tableData;
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
    
    // keyは置換ワードと同じにする
    if (monthStr === targetMonth) {
      monthData = {
        targetMonth: Utilities.formatDate(new Date(`${targetMonth}-01`), 'Asia/Tokyo', 'yyyy年M月'),
        totalWorkDays: data[i][1] || 0,
        summary: data[i][2] || 0,
        incomplete: data[i][3] || 0,
        remarks: data[i][4] || '',
        status: data[i][5] || '',
      };
      break;
    }
  }
  
  // 詳細データを取得
  if (monthData) {
    // 作業内容詳細を取得
    const details = getWorkDetailsByMonthForMakeReport(targetMonth);
    
    // 詳細テーブル用のHTML形式データを作成 (HTML挿入用)
    let detailsTable = '';
    if (details.length > 0) {
      detailsTable = '<table border="1" width="100%">\n';
      detailsTable += '<tr><th>日付</th><th>稼働時間(h)</th><th>業務内容</th><th>ステータス</th></tr>\n';
      
      details.forEach(detail => {
        detailsTable += `<tr>
          <td>${detail.workDate}</td>
          <td>${detail.workTime}</td>
          <td>${detail.workContent}</td>
          <td>${detail.status}</td>
        </tr>\n`;
      });
      
      detailsTable += '</table>';
    } else {
      detailsTable = '詳細データがありません';
    }
    
    // データに追加
    monthData.detailsTable = detailsTable;
    monthData.details = details; // 生のデータも追加
    
    // テキスト形式の詳細リストも作成（HTMLが使えない場合のフォールバック）
    monthData.workDetails = details.map((detail, i) => {
      return `${i+1}. ${detail.workDate} (${detail.workTime}h): ${detail.workContent} [${detail.status}]`;
    }).join('\n');
  }
  
  return monthData;
}

