const GEMINI_KEY = PropertiesService.getScriptProperties()
                    .getProperty("GEMINI_KEY");

const GEMINI_URL =
  "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" +
  GEMINI_KEY;
    

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("AI Tools")
    .addItem("Analyse all rows", "analyseAll")
    .addToUi();
}
function analyseAll() {
  const sheet = SpreadsheetApp.getActiveSheet();

  // Set headers
  sheet.getRange("B1").setValue("Sentiment");
  sheet.getRange("C1").setValue("Summary");

  let row = 2;

  while (true) {
    const text = sheet.getRange(row, 1).getValue();

    if (text === "" || text === null) break;

    const result = analyseText(text);

    sheet.getRange(row, 2).setValue(result.sentiment);
    sheet.getRange(row, 3).setValue(result.summary);

    row++; // move to next row
    Utilities.sleep(500); // avoid rate limits
  }
}
function analyseText(text) {
  const prompt = `Analyse this text. Reply ONLY as JSON with keys 
"sentiment" (Positive/Negative/Neutral) and "summary" (max 20 words).
Text: "${text}"`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  try {
    const res = UrlFetchApp.fetch(GEMINI_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });
    const raw = JSON.parse(res.getContentText());
    const reply = raw.candidates[0].content.parts[0].text
                    .replace(/```json|```/g, "").trim();
    return JSON.parse(reply);
  } catch (e) {
    return { sentiment: "Error", summary: e.message };
  }
}
