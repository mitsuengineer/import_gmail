function importGmailToSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "メールインポート";
  var sheet = spreadsheet.getSheetByName(sheetName);

  // 検索条件を指定
  var subjectText = "年会費"; // 件名を指定
  var startDate = "2024-01-01"; // 開始日を指定
  var endDate = "2024-12-31"; // 終了日を指定

  // 指定のクエリを使用してスレッドを検索（条件がnullでも許容する）
  var query = [
    subjectText ? 'subject:(' + subjectText + ')' : null,
    startDate ? 'after:' + startDate : null,
    endDate ? 'before:' + endDate : null
  ].filter(Boolean).join(' ');

  // Gmailからスレッドを取得
  var threads = GmailApp.search(query);
  var lastRow = sheet.getLastRow();
  var numRows = Math.max(0, lastRow - 1); // numRowsは1以上になるよう調整（1行目をタイトルとして除外）
  var existingThreadIds = numRows ? sheet.getRange(2, 1, numRows, 1).getValues().flat() : [];

  for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      var firstMessage = messages[0];
      var lastMessage = messages[messages.length - 1];
      var threadId = threads[i].getId();  // スレッドIDを取得
      var subject = lastMessage.getSubject();  // 件名を取得
      var lastReplyDate = lastMessage.getDate(); // 最終返信日付を取得
      var from = lastMessage.getFrom(); // 最終返信Fromを取得
      var to = lastMessage.getTo(); //最終返信Toを取得
      var cc = lastMessage.getCc() || '';  //最終返信Ccを取得（Ccがない場合は空文字を設定）
      var firstMessageBody = firstMessage.getPlainBody();  // 最初のメールの本文を取得
      var lastMessageBody = lastMessage.getPlainBody();  // 最終返信本文を取得
      var messageCount = messages.length;  // やり取り数を取得

      // 最終返信者のチェック（最終返信が自分の場合は"弊社", それ以外は"未返信"）
      var status = (from.indexOf("xxx@gmail.com") !== -1) ? "弊社" : "未返信"; // xxxは自分のアドレスをセット

      // 重複チェック
      var rowIndex = existingThreadIds.indexOf(threadId);
      if (rowIndex !== -1) {
          // 既存のデータを上書き
          sheet.getRange(rowIndex + 2, 1, 1, 10).setValues([[threadId, subject, lastReplyDate, from, to, cc, lastMessageBody, firstMessageBody, messageCount, status]]);
      } else {
          // 新しいデータを追加（1行目をタイトル行として保持）
          var nextRow = Math.max(sheet.getLastRow(), 1) + 1; // タイトル行を考慮して次の行を計算
          sheet.getRange(nextRow, 1, 1, 10).setValues([[threadId, subject, lastReplyDate, from, to, cc, lastMessageBody, firstMessageBody, messageCount, status]]);
    }
  }
}