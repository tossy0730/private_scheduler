/**
 * すべての「YYYY年M月」シートを走査して
 * Googleカレンダーに予定を上書き反映するメイン関数
 * （GASトリガーはこの関数を1日おきで設定）
 */
function autoUpdateAllAvailableSchedules() {
  const calendar = CalendarApp.getDefaultCalendar();
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const sheetNameRegex = /^(\d{4})年(\d{1,2})月$/; // 例: 2025年9月

  sheets.forEach(sheet => {
    // 無名やゴミシートを安全にスキップ
    const sheetName = sheet.getName();
    if (!sheetName) return;

    const m = sheetName.match(sheetNameRegex);
    if (!m) return; // 想定形式じゃないシートは無視

    const year = parseInt(m[1], 10);
    const monthZeroBased = parseInt(m[2], 10) - 1; // JSの月は0始まり
    const lastRow = sheet.getLastRow();

    // データがヘッダのみ（2行以下）の場合はスキップ
    if (lastRow <= 2) return;

    // A3:C最終行を取得（A:日付, B:曜日, C:予定）
    const values = sheet.getRange(3, 1, lastRow - 2, 3).getValues();

    values.forEach(row => {
      const day = row[0];        // A: 日付（1,2,3,...）
      const text = String(row[2] || '').trim(); // C: 予定

      // 入力なし or OFFはスキップ（大文字小文字ゆるく）
      if (!day || !text || text.toUpperCase().includes('OFF')) return;

      const date = new Date(year, monthZeroBased, day);

      // `/` 区切りで複数予定に対応
      const chunks = text.split('/').map(s => s.trim()).filter(Boolean);

      chunks.forEach(chunk => {
        try {
          // 例: "10:00-14:00（応援）" から時間を抽出
          const timeMatch = chunk.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
          // タイトルだけに整形（括弧は見た目用に除去）
          const titleOnly = chunk
            .replace(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/, '')
            .replace(/[()（）]/g, '')
            .trim() || '部活';

          // 同日・同タイトルは削除（上書きのため）
          const sameDayEvents = calendar.getEventsForDay(date);
          sameDayEvents.forEach(ev => {
            const t = (ev.getTitle() || '').trim();
            if (t === titleOnly || t === chunk) {
              ev.deleteEvent();
            }
          });

          if (timeMatch) {
            // 時間ありイベント
            const sh = parseInt(timeMatch[1], 10);
            const sm = parseInt(timeMatch[2], 10);
            const eh = parseInt(timeMatch[3], 10);
            const em = parseInt(timeMatch[4], 10);

            const start = new Date(year, monthZeroBased, day, sh, sm);
            const end   = new Date(year, monthZeroBased, day, eh, em);

            // 終了が開始より前/同じにならないようにガード
            if (end <= start) {
              // もし時間表記が変でも、とりあえず1時間イベントにする
              const safeEnd = new Date(start.getTime() + 60 * 60 * 1000);
              calendar.createEvent(titleOnly, start, safeEnd);
            } else {
              calendar.createEvent(titleOnly, start, end);
            }
          } else {
            // 終日イベント
            calendar.createAllDayEvent(titleOnly, date);
          }
        } catch (e) {
          // 1件でエラーっても他は回す
          console.error('Event create error:', e);
        }
      });
    });
  });
}
