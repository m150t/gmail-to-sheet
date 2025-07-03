function extractCrowdWorksEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let threads;
  try {
    threads = GmailApp.search('from:no-reply@crowdworks.jp subject:【クラウドワークス】');
  } catch (error) {
    Logger.log("Gmail検索エラー: " + error.message);
    return;
  }

  const lastRow = sheet.getLastRow();
  const existingIds = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1).getValues().flat()
    : [];

  for (const thread of threads) {
    try {
      const threadId = thread.getId();
      if (existingIds.includes(threadId)) continue;

      const message = thread.getMessages()[0];
      const date = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
      const subject = message.getSubject();
      const rawBody = message.getPlainBody().trim();

      // ヘッダー部分削除（必要なら）
      const cleanedBody = rawBody.replace(
        /={60}\r?\nこのメールはHTMLメールを自動変換して作成しております。\r?\nもし見づらい箇所がある場合は、HTMLメールを見る事ができる\r?\nアドレスに転送するなどしてご覧ください。\r?\n={60}\r?\n?/,
        ''
      ).trim();

      // カテゴリ分類
      let category = '';
      if (subject.includes('相談がありました')) {
        category = 'スカウト';
      } else if (subject.includes('ご応募いただきありがとうございます')) {
        category = '応募';
      } else if (subject.includes('応募・スカウトが辞退されました')) {
        category = '辞退';
      } else {
        category = 'その他';
      }

      // 転記（A:スレッドID, B:カテゴリ, C:日付, D:件名, E:本文）
      sheet.appendRow([threadId, category, date, subject, cleanedBody]);
    } catch (err) {
      Logger.log("処理中エラー: " + err.message);
    }
  }
}
