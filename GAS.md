/**
 * ==========================================
 * 1. 設定値 (Configuration)
 * ==========================================
 */
const PROJECT_ID = 'dummy-id';
const LOCATION = 'us-central1';
const MODEL_ID = 'gemini-2.5-flash';
const MAX_OUTPUT_TOKENS = 8192;

// シート設定
const PROMPTS_SHEET_NAME = 'prompts';
const TARGET_SHEET_NAME = '日報';
const MODE_CELL = 'E1';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AIコーチ')
    .addItem('この行のフィードバックをもらう', 'generateFeedbackForActiveRow')
    .addToUi();
}

/**
 * ==========================================
 * 3. メイン処理
 * ==========================================
 */
function generateFeedbackForActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const activeRow = sheet.getActiveCell().getRow();

  // 実行条件チェック
  if (activeRow < 3) {
    ss.toast('3行目以降のデータ行を選択してください。', '実行不可');
    return;
  }

  // データ取得
  const values = sheet.getRange(activeRow, 2, 1, 3).getValues()[0];
  const dateVal = values[0];
  const contentVal = values[1];
  const planVal = values[2];

  if (!contentVal || !planVal) {
    ui.alert('C列（内容）とD列（明日の予定）が入力されている必要があります。');
    return;
  }

  // 日付フォーマット
  const dateStr = (dateVal instanceof Date) 
    ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd") 
    : dateVal;

  // ▼ プロンプトの取得処理
  let template = getPromptTemplate(ss); 

  // テンプレートの置換
  const promptText = template
    .replace(/{{DATE}}/g, dateStr)
    .replace(/{{CONTENT}}/g, contentVal)
    .replace(/{{PLAN}}/g, planVal);

  try {
    ss.toast('Geminiが思考中です...', '処理中');
    const feedback = callVertexAI(promptText);
    sheet.getRange(activeRow, 5).setValue(feedback);
    ss.toast('フィードバックをE列に書き込みました。', '完了');
  } catch (e) {
    console.error(e);
    ui.alert('エラーが発生しました。\n詳細: ' + e.toString());
  }
}

/**
 * ▼ 選択されたモードに合わせてプロンプトを取得する関数
 */
function getPromptTemplate(ss) {
  const promptsSheet = ss.getSheetByName(PROMPTS_SHEET_NAME);
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME) || ss.getActiveSheet(); // 名前が違う場合はアクティブシート
  
  // 1. 選択されているモードを取得 (E1セル)
  let selectedMode = targetSheet.getRange(MODE_CELL).getValue();
  if (!selectedMode) selectedMode = '標準'; // 空なら標準

  // 2. configシートから全プロンプトを取得
  // A列:モード名, B列:本文
  const lastRow = promptsSheet.getLastRow();
  if (lastRow < 2) throw new Error('configシートにプロンプトが設定されていません。');
  
  const data = promptsSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // [[モード名, 本文], ...]
  
  let template = '';

  // 3. モードに応じたロジック
  if (selectedMode === 'ランダム') {
    // ランダムに1つ選ぶ
    const randomIndex = Math.floor(Math.random() * data.length);
    template = data[randomIndex][1];
    ss.toast(`今日の担当: ${data[randomIndex][0]}`, '抽選結果'); // 誰が選ばれたか通知

  } else {
    // 指定されたモードを探す
    const targetRow = data.find(row => row[0] === selectedMode);
    if (targetRow) {
      template = targetRow[1];
    } else {
      // 見つからなければ1行目(標準)を使う
      template = data[0][1];
      ss.toast(`モード「${selectedMode}」が見つからないため、標準モードで実行します。`, '注意');
    }
  }

  return template;
}

/**
 * ==========================================
 * 4. Vertex AI API 呼び出し関数
 * ==========================================
 */
function callVertexAI(prompt) {
  const url = `https://${LOCATION}-aiplatform.googleapis.com/v1beta1/projects/${PROJECT_ID}/locations/${LOCATION}/publishers/google/models/${MODEL_ID}:generateContent`;

  const payload = {
    "contents": [
      {
        "role": "user",
        "parts": [{ "text": prompt }]
      }
    ],
    "generationConfig": {
      "temperature": 0.7,
      "maxOutputTokens": MAX_OUTPUT_TOKENS
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken()
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error(`Gemini API Error: ${json.error.message} (Code: ${json.error.code})`);
  }

  if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) {
    return json.candidates[0].content.parts[0].text.trim();
  } else {
    throw new Error('Geminiからの応答が空、または想定外の形式です。');
  }
}
