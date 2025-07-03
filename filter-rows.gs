function updateRowVisibility() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();

  // ヘッダー以降の行に対して
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[5]; // F列（ステータス）
    const dateStr = row[2]; // C列（応募日）
    let shouldHide = false;

    // ステータスが "-" なら非表示
    if (status === '-') {
      shouldHide = true;
    }

    // "回答待ち" かつ 1ヶ月以上前なら非表示
    if (status === '回答待ち') {
      const date = new Date(dateStr);
      const oneMonthAgo = new Date(today);
      oneMonthAgo.setMonth(today.getMonth() - 1);
      if (date < oneMonthAgo) {
        shouldHide = true;
      }
    }

    // 行番号はインデックス+1
    const rowIndex = i + 1;
    sheet.showRows(rowIndex); // 一旦表示してから
    if (shouldHide) {
      sheet.hideRows(rowIndex);
    }
  }
}
