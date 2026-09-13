function doGet(e) {
  const key = e.parameter.key;
  const sheet = getDataSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) {
      return ContentService.createTextOutput(rows[i][1])
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.action === 'ocr') {
    return handleOcr(body);
  }

  const key = body.key;
  const value = body.value;
  const sheet = getDataSheet();
  const rows = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([key, value]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Data');
  if (!sheet) {
    sheet = ss.insertSheet('Data');
    sheet.appendRow(['key', 'value']);
  }
  return sheet;
}

// Odczyt paragonu/licznika ze zdjęcia przez Gemini API.
// Klucz API: Script Properties (Project Settings -> Script Properties) -> GEMINI_API_KEY.
function handleOcr(body) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return jsonOut({ ok: false, error: 'Brak klucza GEMINI_API_KEY w ustawieniach skryptu (Project Settings -> Script Properties).' });
  }

  const prompt = body.mode === 'licznik'
    ? 'Jesteś asystentem odczytującym przebieg (drogomierz / licznik kilometrów) samochodu ze zdjęcia deski rozdzielczej. ' +
      'Odczytaj widoczną wartość przebiegu w kilometrach i zwróć WYŁĄCZNIE obiekt JSON w formacie: ' +
      '{"odometerKm": liczba_całkowita_lub_null}. Jeśli przebieg nie jest wyraźnie widoczny, ustaw null.'
    : 'Jesteś asystentem odczytującym dane ze zdjęcia paragonu za paliwo lub wyświetlacza dystrybutora. ' +
      'Odczytaj widoczne wartości i zwróć WYŁĄCZNIE obiekt JSON w formacie: ' +
      '{"liters": liczba_lub_null, "pricePerLiter": liczba_lub_null, "totalCost": liczba_lub_null}. ' +
      'Jeśli wartość nie jest widoczna na zdjęciu, ustaw null. Liczby zawsze z kropką jako separatorem dziesiętnym, bez jednostek i bez spacji.';

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: body.mimeType, data: body.dataBase64 } }
      ]
    }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };

  // "gemini-flash-latest" to ruchomy alias Google, zawsze wskazujący aktualny
  // zalecany model flash -- dzięki temu nie trzeba wracać do tego kodu przy
  // każdej zmianie nazwy/wersji modelu przez Google. Reszta listy to zapasowe
  // nazwy na wypadek gdyby alias przestał działać.
  const models = ['gemini-flash-latest', 'gemini-3.6-flash'];
  const cached = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL_OK');
  if (cached && models.indexOf(cached) === -1) models.unshift(cached);

  let lastError = 'nieznany błąd';
  for (let i = 0; i < models.length; i++) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent?key=' + apiKey;
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const status = res.getResponseCode();
      const resJson = JSON.parse(res.getContentText());
      if (status !== 200) {
        lastError = 'model ' + models[i] + ': ' + (resJson.error ? resJson.error.message : status);
        continue;
      }
      PropertiesService.getScriptProperties().setProperty('GEMINI_MODEL_OK', models[i]);
      const text = resJson.candidates[0].content.parts[0].text;
      const data = JSON.parse(text);
      return jsonOut({ ok: true, data: data, model: models[i] });
    } catch (err) {
      lastError = 'model ' + models[i] + ': ' + err.message;
    }
  }
  return jsonOut({ ok: false, error: 'Gemini API: ' + lastError });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
