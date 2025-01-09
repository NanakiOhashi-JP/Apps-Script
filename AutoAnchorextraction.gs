function fetchHrefForAnchorText() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const url = "https://zero2one.jp/ai-word/?srsltid=AfmBOoqyJL8vVhHU8SjLlA0pmmdjsC2rCdEHuwf5ExgrB_cyjmoOXUj2"; // 対象URL
  const html = UrlFetchApp.fetch(url).getContentText(); // WebページのHTMLを取得

  const lastRow = sheet.getLastRow();
  const anchorTexts = sheet.getRange(2, 2, lastRow - 1).getValues(); // B列のアンカーテキストを取得
  const results = [];

  anchorTexts.forEach(([text]) => {
    if (text) {
      const href = findHrefByAnchorText(html, text);
      results.push([href || "該当なし"]); // hrefが見つからなければ"該当なし"と表示
    } else {
      results.push(["入力なし"]);
    }
  });

  sheet.getRange(2, 5, results.length).setValues(results); // E列に結果を出力
}

function findHrefByAnchorText(html, anchorText) {
  const regex = new RegExp(`<a[^>]*href="([^"]*)"[^>]*>${anchorText}</a>`, "gi"); // アンカーテキストに一致するhrefを抽出
  const match = regex.exec(html);
  return match ? match[1] : null; // hrefが見つかれば返す
}
