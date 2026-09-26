// Musi być identyczna wartość początkowa jak w index.html (PIN_INITIAL) —
// jeśli PIN nigdy nie był zmieniony przez "Zmień PIN", to jest ta,
// którą trzeba wpisać. Po pierwszej zmianie prawdziwy PIN żyje tylko
// w PropertiesService, ta stała już nie ma znaczenia.
const PIN_INITIAL = '1000499156';
const PIN_RESET_EMAIL = 'heatcoolfulawkawro@gmail.com';
// Zwracany zamiast danych z fetch(GET), gdy PIN się nie zgadza — pusty
// string ('') już oznacza "brak takiego klucza", więc potrzebny jest
// osobny, jednoznaczny sygnał, którego żadna prawdziwa wartość nigdy
// nie przyjmie.
const AUTH_FAIL_TEXT = '__BRAK_AUTORYZACJI__';

// ---------- Sync PIN-u z siostrzanymi appkami (ten sam PF/admin) ----------
// Żeby dołożyć kolejną appkę do tej samej "rodziny" jednego kodu:
//   1) w NOWEJ appce wklej dokładnie ten sam blok kodu (SIBLING_URLS,
//      bootstrapSyncSecret, syncPinPush, pushPinToSiblings) i dopisz wywołanie
//      pushPinToSiblings(newPin) na końcu jej confirmPinReset — patrz niżej.
//   2) do SIBLING_URLS TEJ appki i wszystkich pozostałych już istniejących
//      dopisz URL nowej appki (i dopisz URL-e istniejących do listy nowej).
//   3) zbootstrapuj w nowej appce TEN SAM sekret co reszta rodziny (jednym
//      POST-em z action:'bootstrap_sync_secret' — działa tylko raz, dopóki
//      SYNC_SECRET jest puste).
const SIBLING_URLS = [
  'https://script.google.com/macros/s/AKfycbz3-nc9P2jTv3pX2_aiP6Ne7A67QXtZHObP53BU3GNMIjgrThQSJtfaOCnBbGSGSRQI/exec' // Waga PF
];

function bootstrapSyncSecret(secret) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SYNC_SECRET')) return jsonOut({ ok: false, error: 'Sekret już ustawiony' });
  if (!secret || String(secret).length < 20) return jsonOut({ ok: false, error: 'Za krótki sekret' });
  props.setProperty('SYNC_SECRET', String(secret));
  return jsonOut({ ok: true });
}

// Odbiór PIN-u z siostrzanej appki — NIE rozsyła dalej (jeden przeskok,
// żeby appki nie wołały się w kółko).
function syncPinPush(secret, newPin) {
  const real = PropertiesService.getScriptProperties().getProperty('SYNC_SECRET');
  if (!real || String(secret) !== real) return jsonOut({ ok: false, error: 'Brak autoryzacji' });
  if (!/^\d{4,12}$/.test(String(newPin))) return jsonOut({ ok: false, error: 'Zły format PIN' });
  PropertiesService.getScriptProperties().setProperty('APP_PIN', String(newPin));
  return jsonOut({ ok: true });
}

// Diagnostyka: sprawdza, czy ta appka faktycznie dobija się (POST, nie GET
// po przekierowaniu) do każdej appki z SIBLING_URLS i czy sekret się zgadza —
// bez dotykania PIN-u. sync_ping po drugiej stronie tylko potwierdza sekret.
function syncSelftest() {
  const secret = PropertiesService.getScriptProperties().getProperty('SYNC_SECRET');
  if (!secret) return jsonOut({ ok: false, error: 'Brak SYNC_SECRET — najpierw bootstrap' });
  const results = SIBLING_URLS.map(function (url) {
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'text/plain',
        payload: JSON.stringify({ action: 'sync_ping', secret: secret }),
        muteHttpExceptions: true,
        followRedirects: true
      });
      return { url: url, status: res.getResponseCode(), body: res.getContentText().slice(0, 300) };
    } catch (e) {
      return { url: url, error: e.message };
    }
  });
  return jsonOut({ ok: true, results: results });
}

// Wywoływane PO stronie appki, w której PIN faktycznie się zmienił —
// rozsyła nowy PIN do sióstr. Najlepszego wysiłku: appka, która akurat nie
// odpowie, dogoni przy najbliższym auth-fail (pokaże błąd, pójdzie reset mailem).
function pushPinToSiblings(newPin) {
  const secret = PropertiesService.getScriptProperties().getProperty('SYNC_SECRET');
  if (!secret) return;
  SIBLING_URLS.forEach(function (url) {
    try {
      UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'text/plain',
        payload: JSON.stringify({ action: 'sync_pin_push', secret: secret, newPin: newPin }),
        muteHttpExceptions: true
      });
    } catch (e) { /* best-effort — patrz komentarz wyżej */ }
  });
}

function currentPin() {
  return PropertiesService.getScriptProperties().getProperty('APP_PIN') || PIN_INITIAL;
}

function doGet(e) {
  if (String(e.parameter.pin) !== currentPin()) {
    return ContentService.createTextOutput(AUTH_FAIL_TEXT).setMimeType(ContentService.MimeType.JSON);
  }
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

  if (body.action === 'request_pin_reset') return requestPinReset();
  if (body.action === 'confirm_pin_reset') return confirmPinReset(body.code, body.newPin);
  if (body.action === 'sync_pin_push') return syncPinPush(body.secret, body.newPin);
  if (body.action === 'bootstrap_sync_secret') return bootstrapSyncSecret(body.secret);
  if (body.action === 'sync_selftest') return syncSelftest();
  if (body.action === 'sync_ping') {
    const real = PropertiesService.getScriptProperties().getProperty('SYNC_SECRET');
    return jsonOut({ ok: !!real && String(body.secret) === real });
  }

  if (String(body.pin) !== currentPin()) {
    return jsonOut({ ok: false, error: 'Brak autoryzacji' });
  }

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

  // Uwaga dot. "licznik": schemat JSON wymusza natychmiastową odpowiedź bez
  // miejsca na "przemyślenie" -- dlatego każemy modelowi najpierw przepisać
  // cyfry pojedynczo do "digitsRead" (odpowiednik chain-of-thought przy
  // sztywnym JSON), a dopiero potem złożyć to w odometerKm. To wyraźnie
  // poprawia trafność odczytu wyświetlaczy 7-segmentowych / mechanicznych
  // bębenków w porównaniu do proszenia od razu o gotową liczbę.
  const prompt = body.mode === 'licznik'
    ? 'Jesteś precyzyjnym asystentem odczytującym przebieg (drogomierz / odometer) samochodu ze zdjęcia zestawu wskaźników. Zdjęcie jest zwykle mocno przybliżone i pokazuje głównie sam wyświetlacz licznika, bez szerszego kontekstu deski rozdzielczej. Wykonaj kolejno: ' +
      '1) Znajdź licznik CAŁKOWITEGO przebiegu pojazdu (ODO) -- to zwykle największa, najbardziej wyeksponowana liczba na wyświetlaczu, 5-7 cyfr, czasem z jednostką "km" tuż obok. NIE bierz licznika TRASY (TRIP / TRIP A / TRIP B / dzienny) -- to inna, resetowalna wartość, zwykle mniejsza, często z miejscem po przecinku (np. 234.5) i/lub etykietą "trip"/"A"/"B" obok; jeśli widzisz oba liczniki na zdjęciu, wybierz ten BEZ takiej etykiety. ' +
      '2) Wyświetlacz może być cyfrowym LCD (segmentowym) albo mechanicznym bębenkiem z cyframi na wałkach -- w obu przypadkach czytaj cyfry pojedynczo, od lewej do prawej, zwracając uwagę na cyfry łatwe do pomylenia (3/8, 5/6, 1/7), niedoświetlone/częściowo zasłonięte segmenty oraz na to, że ostatni bębenek/cyfra licznika mechanicznego bywa "w trakcie obrotu" między dwiema wartościami -- w takim wypadku przyjmij cyfrę, która zajmuje większą część pola. ' +
      '3) Zwróć WYŁĄCZNIE obiekt JSON w formacie: {"digitsRead": "cyfry oddzielone spacją np. 5 0 0 5 8", "odometerKm": liczba_całkowita_lub_null}. Jeśli przebiegu nie da się jednoznacznie odczytać (np. całkowicie zasłonięty, nieostry, brak licznika na zdjęciu), ustaw odometerKm na null zamiast zgadywać wartość.'
    : 'Jesteś precyzyjnym asystentem odczytującym dane ze zdjęcia paragonu za paliwo lub wyświetlacza dystrybutora. Wyświetlacz dystrybutora to zwykle cyfry LCD/LED -- czytaj je uważnie cyfra po cyfrze, zwracając uwagę na cyfry łatwe do pomylenia (3/8, 5/6, 1/7) i na miejsca po przecinku. ' +
      'Zwróć WYŁĄCZNIE obiekt JSON w formacie: ' +
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

// ---------- PIN: zmiana / przypomnienie przez kod z maila ----------
// Bez tokenu — to jedyna para akcji dostępna komuś, kto NIE zna aktualnego
// PIN-u (inaczej "zapomniałem PIN-u" nie dałoby się obsłużyć). Limit czasowy
// między żądaniami to jedyna ochrona przed zasypaniem skrzynki e-mail.
function requestPinReset() {
  const props = PropertiesService.getScriptProperties();
  const lastReq = Number(props.getProperty('PIN_RESET_LAST_REQ') || 0);
  if (Date.now() - lastReq < 2 * 60 * 1000) {
    return jsonOut({ ok: false, error: 'Poczekaj chwilę i spróbuj ponownie.' });
  }
  const code = String(Math.floor(100000 + Math.random() * 900000));
  props.setProperty('PIN_RESET_CODE', code);
  props.setProperty('PIN_RESET_EXPIRES', String(Date.now() + 10 * 60 * 1000));
  props.setProperty('PIN_RESET_LAST_REQ', String(Date.now()));
  MailApp.sendEmail(PIN_RESET_EMAIL, 'Kod do zmiany PIN — Paliwo PF', 'Twój kod do zmiany PIN: ' + code + '\n\nWażny 10 minut. Jeśli to nie Ty, zignoruj tę wiadomość.');
  return jsonOut({ ok: true });
}

function confirmPinReset(code, newPin) {
  const props = PropertiesService.getScriptProperties();
  const storedCode = props.getProperty('PIN_RESET_CODE');
  const expires = Number(props.getProperty('PIN_RESET_EXPIRES') || 0);
  if (!storedCode || String(code) !== storedCode) return jsonOut({ ok: false, error: 'Nieprawidłowy kod' });
  if (Date.now() > expires) return jsonOut({ ok: false, error: 'Kod wygasł — poproś o nowy' });
  if (!/^\d{4,12}$/.test(String(newPin))) return jsonOut({ ok: false, error: 'PIN musi mieć od 4 do 12 cyfr' });
  props.setProperty('APP_PIN', String(newPin));
  props.deleteProperty('PIN_RESET_CODE');
  props.deleteProperty('PIN_RESET_EXPIRES');
  pushPinToSiblings(String(newPin));
  return jsonOut({ ok: true });
}
