// ユーティリティ関数のインクルード
//#include "utils/configUtils.gs"

function importWorkingDaysForReport() {
  // 環境設定値を取得
  const config = getConfigFromProperties();
  if (!config) return; // 設定が不足している場合は処理を中断
  
  const calendarId = config.calendarId;
  const WORKING_START_DATE = config.workingStartDate;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('稼働一覧表');
  if (!sheet) {
    console.error('「稼働一覧表」シートが見つかりません。');
    return;
  }
  const calendar = CalendarApp.getCalendarById(calendarId);
  const lastRow = sheet.getLastRow();

  // === 既存データをマップに読み込む（イベントIDをキーにする）===
  const existingData = {};
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // 2行目から5列分（A〜E列）を取得
    values.forEach((row, index) => {
      const date = row[0];
      const eventId = row[4];
      if (eventId) {
        existingData[eventId] = {
          date: date,
          durationHours: row[1],
          description: row[2],
          status: row[3],
          rowNumber: index + 2
        };
      }
    });
  }

  // === 取得対象期間を設定 ===
  const today = new Date();
  const end = new Date(today.getFullYear() + 1, today.getMonth(), today.getDate()); // 今日から1年後

  const events = calendar.getEvents(WORKING_START_DATE, end); // ← ここでWORKING_START_DATEを使用！

  // === 「稼働日」をタイトルに含む予定だけフィルタリング ===
  const workingDayEvents = events.filter(event => event.getTitle().includes('稼働日'));

  const usedEventIds = new Set(); // 処理したイベントIDを保存する（後で削除判定に使う）

  // === カレンダー予定を1件ずつ処理 ===
  workingDayEvents.forEach(event => {
    const eventId = event.getId();
    const startTime = event.getStartTime();
    const endTimeEvent = event.getEndTime();
    const durationMs = endTimeEvent - startTime;
    const durationHours = Math.round((durationMs / (1000 * 60 * 60)) * 10) / 10;
    let description = event.getDescription() || '';

    // 説明が改行のみの場合は空文字に置き換える
    if (description.trim() === '') {
      description = '';
    }

    // 日付（開始日のみ）
    const dateOnly = new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate());

    usedEventIds.add(eventId); // このイベントIDは有効なので記録しておく

    if (existingData[eventId]) {
      const existing = existingData[eventId];
      // === 日付・稼働時間・業務内容のいずれかに違いがあれば、上書き更新 ===
      if (existing.date.getTime() !== dateOnly.getTime() ||
          existing.durationHours !== durationHours ||
          existing.description !== description) {
        sheet.getRange(existing.rowNumber, 1, 1, 5).setValues([
          [dateOnly, durationHours, description, existing.status, eventId]
        ]);
      }
    } else {
      // === 新しいイベントなら、スプレッドシートに追加 ===
      sheet.appendRow([
        dateOnly,
        durationHours,
        description,
        '', // ステータスは空欄
        eventId
      ]);
    }
  });

  // === 削除対象（カレンダーに存在しなくなったイベント）を検出・削除 ===
  const allEventIds = Object.keys(existingData);
  const deletedEventIds = allEventIds.filter(id => !usedEventIds.has(id));

  // 削除は下から順番に行う（行番号がズレないように）
  deletedEventIds.sort((a, b) => existingData[b].rowNumber - existingData[a].rowNumber);

  deletedEventIds.forEach(eventId => {
    const rowNumber = existingData[eventId].rowNumber;
    sheet.deleteRow(rowNumber);
  });

  // === 最後に、日付順（昇順）に並べ替え ===
  const newLastRow = sheet.getLastRow();
  if (newLastRow > 1) {
    sheet.getRange(2, 1, newLastRow - 1, 5).sort({column: 1, ascending: true});
  }
}
