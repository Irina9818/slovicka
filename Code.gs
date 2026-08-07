/**
 * SLOVÍČKA - backend na Google Apps Script
 * Sloupce v listu "Slovicka": id | word | translation | example | box | addedDate | nextReview | lastReviewed | timesCorrect | timesWrong
 */

const SHEET_NAME = 'Slovicka';
const BOX_INTERVALS = {1: 1, 2: 2, 3: 4, 4: 7, 5: 14}; // dny do dalšího opakování

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'due') return jsonOutput(getDueWords());
  if (action === 'all') return jsonOutput(getAllWords());
  if (action === 'add') return jsonOutput(addWord({ word: e.parameter.word }));
  if (action === 'review') return jsonOutput(reviewWord({ id: e.parameter.id, correct: e.parameter.correct === 'true' }));
  return jsonOutput({error: 'unknown action: ' + action});
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === 'add') return jsonOutput(addWord(data));
  if (data.action === 'review') return jsonOutput(reviewWord(data));
  return jsonOutput({error: 'unknown action: ' + data.action});
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('List "' + SHEET_NAME + '" neexistuje. Vytvoř ho a přidej hlavičku.');
  return sheet;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function addWord(data) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const today = new Date();

  const generated = generateTranslationAndExample(data.word);

  if (!generated.valid) {
    return { status: 'error', message: generated.reason || 'AI nerozpoznala platné slovo ani frázi.' };
  }

  // id, word (EN), translation (CS), example, box, addedDate, nextReview, lastReviewed, timesCorrect, timesWrong
  sheet.appendRow([id, generated.english, generated.czech, generated.example, 1, today, today, '', 0, 0]);
  return { status: 'ok', id: id, word: generated.english, translation: generated.czech, example: generated.example };
}

/**
 * Zavolá Claude (Anthropic API). Rozpozná, jestli je vstup anglicky nebo česky,
 * ověří, že jde o reálné slovo/frázi, a vrátí anglický tvar + český překlad + příklad.
 * Vyžaduje Script Property "ANTHROPIC_API_KEY" (Project Settings → Script Properties v Apps Scriptu).
 */
function generateTranslationAndExample(word) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    throw new Error('Chybí ANTHROPIC_API_KEY ve Script Properties (Nastavení projektu → Vlastnosti skriptu).');
  }

  const prompt = 'Text od uživatele (může být anglicky nebo česky, nebo to může být nesmysl/překlep): "' + word + '"\n\n' +
    'Vrať POUZE syrový JSON, bez markdown formátování, bez ```json bloků, bez jakéhokoli dalšího textu před nebo za JSON, přesně v tomto formátu:\n' +
    '{"valid": true, "english": "anglický tvar slova nebo fráze", "czech": "český překlad", "example": "přirozená anglická příkladová věta používající to slovo v pracovním/běžném kontextu", "reason": ""}\n\n' +
    'Pravidla:\n' +
    '- Pokud je vstup platné anglické NEBO české slovo/fráze, nastav "valid": true a vyplň "english" i "czech" (bez ohledu na to, v jakém jazyce byl vstup napsaný) a "example".\n' +
    '- Pokud vstup není rozpoznatelné slovo ani fráze v žádném z těchto dvou jazyků (překlep, náhodné znaky, nesmysl), nastav "valid": false, "english" a "czech" nech jako prázdný řetězec, "example" nech prázdný, a do "reason" napiš krátké vysvětlení česky, proč to nebylo přijato.\n' +
    '- Pokud má slovo víc významů, vyber ten nejběžnější v každodenní pracovní komunikaci.';

  const payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 300,
    messages: [{ role: 'user', content: prompt }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error('Anthropic API chyba: ' + json.error.message);
  }

  const textBlock = json.content.filter(function (b) { return b.type === 'text'; })[0];
  if (!textBlock) throw new Error('AI nevrátila žádný text.');

  // Odstranit případné markdown code fence bloky (```json ... ```), které AI někdy přidá navíc
  const cleaned = textBlock.text.trim().replace(/^```json\s*/i, '').replace(/^```\s*/, '').replace(/```\s*$/, '').trim();

  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    throw new Error('Nepodařilo se rozparsovat odpověď AI: ' + textBlock.text);
  }

  return parsed;
}

function getDueWords() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const due = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue; // prázdný řádek
    const nextReview = row[6] ? new Date(row[6]) : today;
    nextReview.setHours(0, 0, 0, 0);
    if (nextReview <= today) {
      due.push(rowToObj(row, i + 1));
    }
  }
  return due;
}

function getAllWords() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push(rowToObj(rows[i], i + 1));
  }
  return result;
}

function rowToObj(row, rowIndex) {
  return {
    rowIndex: rowIndex,
    id: row[0], word: row[1], translation: row[2], example: row[3],
    box: row[4], addedDate: row[5], nextReview: row[6], lastReviewed: row[7],
    timesCorrect: row[8], timesWrong: row[9]
  };
}

function reviewWord(data) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      let box = Number(rows[i][4]) || 1;
      let correct = Number(rows[i][8]) || 0;
      let wrong = Number(rows[i][9]) || 0;

      if (data.correct) {
        box = Math.min(box + 1, 5);
        correct += 1;
      } else {
        box = 1;
        wrong += 1;
      }

      const today = new Date();
      const next = new Date();
      next.setDate(today.getDate() + BOX_INTERVALS[box]);

      const rowNum = i + 1;
      sheet.getRange(rowNum, 5).setValue(box);       // E: box
      sheet.getRange(rowNum, 7).setValue(next);       // G: nextReview
      sheet.getRange(rowNum, 8).setValue(today);      // H: lastReviewed
      sheet.getRange(rowNum, 9).setValue(correct);    // I: timesCorrect
      sheet.getRange(rowNum, 10).setValue(wrong);     // J: timesWrong

      return {status: 'ok', newBox: box, nextReview: next};
    }
  }
  return {status: 'error', message: 'word not found: ' + data.id};
}
