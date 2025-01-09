function fetchHrefForAnchorText() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const url = "https://zero2one.jp/ai-word/?srsltid=AfmBOoqyJL8vVhHU8SjLlA0pmmdjsC2rCdEHuwf5ExgrB_cyjmoOXUj2"; // Target URL
  
  let html;
  try {
    // Fetch the web page with a custom User-Agent header
    const options = {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
      }
    };
    html = UrlFetchApp.fetch(url, options).getContentText();
  } catch (e) {
    Logger.log("Error fetching the webpage: " + e.message);
    html = null; // Handle the error gracefully
  }

  if (!html) {
    Logger.log("Failed to retrieve the webpage.");
    return;
  }

  const lastRow = sheet.getLastRow();
  const anchorTexts = sheet.getRange(2, 2, lastRow - 1).getValues(); // Get anchor text from column B
  const results = [];

  anchorTexts.forEach(([text]) => {
    if (text) {
      const href = findHrefByAnchorText(html, text);
      results.push([href || "該当なし"]); // Display "該当なし" if no match is found
    } else {
      results.push(["入力なし"]); // Display "入力なし" if anchor text is empty
    }
  });

  sheet.getRange(2, 5, results.length).setValues(results); // Output results to column E
}

function findHrefByAnchorText(html, anchorText) {
  // Regex to find the href value for the given anchor text
  const regex = new RegExp(`<a[^>]*href="([^"]*)"[^>]*>\\s*${anchorText}\\s*</a>`, "gi");
  const match = regex.exec(html);
  return match ? match[1] : null; // Return href if found
}
