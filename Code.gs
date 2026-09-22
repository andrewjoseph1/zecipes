// ============================================================
// ZECIPES — Apps Script Backend v3
// ============================================================
// CHANGELOG from v2:
//   - Request parsing: doPost now reads fields from e.parameter AND from
//     e.postData.contents (JSON, text/plain JSON, urlencoded, multipart).
//     Large pasted text / photo payloads no longer depend on Apps Script
//     copying form fields into e.parameter.
//   - Text import preserves line structure (newlines are kept, only
//     horizontal whitespace is collapsed). Cap raised to 40k chars.
//   - URL import: HTTP status is checked, bot-challenge and thin pages are
//     detected, and failures return { code: 'BLOCKED' | 'EMPTY' | ... ,
//     fallback: 'paste' } so the frontend can offer paste-import.
//   - JSON-LD recipes are rendered to compact readable text instead of raw
//     JSON (reviews, ratings, images stripped) so instructions are never
//     truncated by the size cap.
//   - HTML fallback keeps block structure (block tags -> newlines, entities
//     decoded, nav/footer/aside removed).
//   - Claude call: single shared function, structured JSON output
//     (output_config.format) on models that support it, one retry on
//     429/5xx/529, API errors surfaced with their real message, recipe
//     fields normalised (arrays -> lines, numbering stripped, category
//     snapped to the allowed list).
//   - Model is configurable via a CLAUDE_MODEL row in the config sheet
//     (default: claude-sonnet-5).
//   - doGet/doPost are wrapped so an exception returns JSON, never an
//     HTML error page the frontend can't parse.
// ============================================================

const SHEET_NAME     = 'recipes';
const MEALPLAN_SHEET = 'mealplan';
const LOG_SHEET      = 'logs';
const CONFIG_SHEET   = 'config';
const SS = SpreadsheetApp.getActiveSpreadsheet();

const DEFAULT_MODEL   = 'claude-sonnet-5';
const TEXT_CHAR_LIMIT = 40000;
const CATEGORIES = ['Breakfast', 'Lunch', 'Dinner', 'Soup', 'Salad', 'Dessert', 'Snack', 'Bread', 'Sauce', 'Drink', 'Side'];

// Models that accept output_config.format (structured JSON output).
// Anything else falls back to prompt-only JSON with tolerant parsing.
const STRUCTURED_OUTPUT_MODELS = /^claude-(sonnet-5|opus-5|opus-4-8|haiku-4-5|fable-5)/;

// ── Helpers ───────────────────────────────────────────────

function getSheet() {
  return SS.getSheetByName(SHEET_NAME);
}

function getMealplanSheet() {
  let sheet = SS.getSheetByName(MEALPLAN_SHEET);
  if (!sheet) {
    sheet = SS.insertSheet(MEALPLAN_SHEET);
    sheet.appendRow(['date', 'recipe_id', 'meal_slot']);
  }
  return sheet;
}

function getLogSheet() {
  let sheet = SS.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = SS.insertSheet(LOG_SHEET);
    sheet.appendRow(['timestamp', 'level', 'action', 'message']);
  }
  return sheet;
}

function logEvent(level, action, message) {
  try {
    const sheet = getLogSheet();
    sheet.appendRow([new Date().toISOString(), level, action, String(message).slice(0, 500)]);
    const rows = sheet.getLastRow();
    if (rows > 201) sheet.deleteRows(2, rows - 201);
  } catch (e) { /* silent */ }
}

function generateId() {
  return 'rec_' + Utilities.getUuid().slice(0, 8);
}

function rowToRecipe(row) {
  return {
    id:           row[0],
    title:        row[1],
    source_url:   row[2],
    category:     row[3],
    servings:     row[4],
    prep_time:    row[5],
    cook_time:    row[6],
    ingredients:  row[7],
    instructions: row[8],
    notes:        row[9]
  };
}

function getAllRecipes() {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(rowToRecipe).filter(r => r.id !== '');
}

function getRecipeById(id) {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return rowToRecipe(data[i]);
  }
  return null;
}

function findRowById(id) {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return null;
}

function getConfig(key, fallback) {
  const sheet = SS.getSheetByName(CONFIG_SHEET);
  if (!sheet) return fallback === undefined ? null : fallback;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      const v = String(data[i][1]).trim();
      return v === '' ? (fallback === undefined ? null : fallback) : v;
    }
  }
  return fallback === undefined ? null : fallback;
}

// ── Request parsing ──────────────────────────────────────
//
// The frontend sends small actions as application/x-www-form-urlencoded
// (Apps Script copies those into e.parameter) and large payloads as a
// text/plain body containing JSON (keeps newlines and structure intact and
// avoids a CORS preflight). Older builds sent multipart FormData. All three
// are merged here so handlers can just read `p.text`, `p.images`, etc.

function parseRequest(e) {
  const p = {};
  if (e && e.parameter) Object.keys(e.parameter).forEach(k => p[k] = e.parameter[k]);

  const pd = e && e.postData;
  if (!pd || !pd.contents) return p;

  const type = String(pd.type || '').toLowerCase();
  const body = pd.contents;

  // 1. JSON body (application/json or text/plain carrying JSON)
  const trimmed = body.trim();
  if (trimmed.charAt(0) === '{' && (type.indexOf('json') !== -1 || type.indexOf('text/plain') !== -1 || type === '')) {
    try {
      const obj = JSON.parse(trimmed);
      Object.keys(obj).forEach(k => { if (p[k] === undefined) p[k] = obj[k]; });
      return p;
    } catch (err) { /* not JSON after all — fall through */ }
  }

  // 2. urlencoded body — usually already in e.parameter, but be safe
  if (type.indexOf('x-www-form-urlencoded') !== -1) {
    body.split('&').forEach(pair => {
      if (!pair) return;
      const idx = pair.indexOf('=');
      const k = decodeURIComponent((idx === -1 ? pair : pair.slice(0, idx)).replace(/\+/g, ' '));
      const v = idx === -1 ? '' : decodeURIComponent(pair.slice(idx + 1).replace(/\+/g, ' '));
      if (p[k] === undefined) p[k] = v;
    });
    return p;
  }

  // 3. multipart/form-data (legacy FormData clients) — text fields only
  if (type.indexOf('multipart/form-data') !== -1 || body.slice(0, 2) === '--') {
    const firstLine = body.split(/\r?\n/, 1)[0];
    if (firstLine && firstLine.slice(0, 2) === '--') {
      const boundary = firstLine.trim();
      body.split(boundary).forEach(part => {
        const m = part.match(/name="([^"]+)"/);
        if (!m) return;
        const sep = part.indexOf('\r\n\r\n') !== -1 ? '\r\n\r\n' : '\n\n';
        const idx = part.indexOf(sep);
        if (idx === -1) return;
        let value = part.slice(idx + sep.length);
        value = value.replace(/\r?\n--\s*$/, '').replace(/\r?\n$/, '');
        if (p[m[1]] === undefined) p[m[1]] = value;
      });
    }
  }
  return p;
}

// ── CRUD — Recipes ───────────────────────────────────────

function addRecipe(params) {
  const sheet = getSheet();
  const id    = generateId();
  sheet.appendRow([
    id,
    params.title        || '',
    params.source_url   || '',
    params.category     || '',
    params.servings     || '',
    params.prep_time    || '',
    params.cook_time    || '',
    params.ingredients  || '',
    params.instructions || '',
    params.notes        || ''
  ]);
  logEvent('INFO', 'add', 'Added: ' + (params.title || id));
  return { success: true, id };
}

function updateRecipe(params) {
  const sheet  = getSheet();
  const rowNum = findRowById(params.id);
  if (!rowNum) return { success: false, error: 'Recipe not found' };
  sheet.getRange(rowNum, 1, 1, 10).setValues([[
    params.id,
    params.title        || '',
    params.source_url   || '',
    params.category     || '',
    params.servings     || '',
    params.prep_time    || '',
    params.cook_time    || '',
    params.ingredients  || '',
    params.instructions || '',
    params.notes        || ''
  ]]);
  logEvent('INFO', 'update', 'Updated: ' + (params.title || params.id));
  return { success: true };
}

function deleteRecipe(id) {
  const sheet  = getSheet();
  const rowNum = findRowById(id);
  if (!rowNum) return { success: false, error: 'Recipe not found' };
  const title = sheet.getRange(rowNum, 2).getValue();
  sheet.deleteRow(rowNum);
  removeMealplanByRecipe(id);
  logEvent('INFO', 'delete', 'Deleted: ' + (title || id));
  return { success: true };
}

// ── CRUD — Meal Plan ─────────────────────────────────────

function getMealplan(startDate, endDate) {
  const sheet = getMealplanSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1)
    .filter(r => r[0] !== '' && r[0] >= startDate && r[0] <= endDate)
    .map(r => ({ date: r[0], recipe_id: r[1], meal_slot: r[2] || 'dinner' }));
}

function setMealplan(date, recipeId, mealSlot) {
  const sheet = getMealplanSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(date) && (data[i][2] || 'dinner') === mealSlot) {
      sheet.getRange(i + 1, 2).setValue(recipeId);
      return { success: true };
    }
  }
  sheet.appendRow([date, recipeId, mealSlot]);
  return { success: true };
}

function removeMealplan(date, mealSlot) {
  const sheet = getMealplanSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(date) && (data[i][2] || 'dinner') === mealSlot) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Entry not found' };
}

function removeMealplanByRecipe(recipeId) {
  const sheet = getMealplanSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === recipeId) sheet.deleteRow(i + 1);
  }
}

function getShoppingList(startDate, endDate) {
  const plan    = getMealplan(startDate, endDate);
  const recipes = getAllRecipes();
  const recipeMap = {};
  recipes.forEach(r => recipeMap[r.id] = r);

  const items = [];
  const usedRecipes = [];
  plan.forEach(p => {
    const r = recipeMap[p.recipe_id];
    if (!r) return;
    usedRecipes.push(r.title);
    (r.ingredients || '').split('\n').filter(Boolean).forEach(ing => {
      items.push(ing.trim());
    });
  });
  return { success: true, items, recipes: usedRecipes };
}

// ── Extraction prompt ─────────────────────────────────────

function recipePrompt() {
  return `Extract the recipe and return ONLY a JSON object with these fields:

- title (string)
- category (string — pick exactly ONE from: ${CATEGORIES.join(', ')}. Pick the closest match.)
- servings (string, e.g. "4" or "serves 4-6")
- prep_time (string, e.g. "15 min")
- cook_time (string, e.g. "30 min")
- ingredients (string, one ingredient per line, format each as: {quantity}{unit} {ingredient}, {prep} — e.g. "200g plain flour" or "2 cloves garlic, minced". Be consistent: always lead with the quantity. If the source groups ingredients under headings like "For the sauce", keep the heading as its own line ending with a colon.)
- instructions (string, one step per line, no numbering — CRITICAL: every step must be fully self-contained with specific quantities and measurements inline so the cook never needs to refer back to the ingredient list. Write "Add 200g flour and 2 tsp baking powder" not "Add the flour and baking powder". Merge trivially short steps like "Season with salt" into the preceding step when it makes sense.)
- notes (string, any tips, substitutions or variations mentioned, or empty string)

IMPORTANT:
- The ingredients list must contain ALL ingredients with full quantities.
- The instructions must ALSO repeat specific quantities inline. Both fields must be independently complete.
- The text may contain ads, navigation, comments, or unrelated content. Ignore everything that is not part of the recipe itself.
- If you find multiple recipes, extract only the primary/featured one.
- Keep the recipe's original language.
- Return ONLY the JSON object, no explanation, no markdown fences.`;
}

function imagePrompt() {
  return recipePrompt() + `

Additional instructions for image extraction:
- If the image is partially cut off or some text is unclear, extract what you can confidently read and mark uncertain parts with [?].
- Do not hallucinate ingredients or quantities you cannot clearly see.
- If multiple pages/images are provided, treat them as parts of the same recipe.`;
}

function recipeSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['title', 'category', 'servings', 'prep_time', 'cook_time', 'ingredients', 'instructions', 'notes'],
    properties: {
      title:        { type: 'string' },
      category:     { type: 'string', enum: CATEGORIES },
      servings:     { type: 'string' },
      prep_time:    { type: 'string' },
      cook_time:    { type: 'string' },
      ingredients:  { type: 'string', description: 'One ingredient per line, separated by newline characters.' },
      instructions: { type: 'string', description: 'One step per line, separated by newline characters, no numbering.' },
      notes:        { type: 'string' }
    }
  };
}

// ── Text normalisation ────────────────────────────────────

// Keeps line structure: collapses runs of spaces/tabs, trims each line,
// collapses 3+ blank lines to one blank line. Newlines are what tell Claude
// where one ingredient ends and the next begins.
function normalizeText(text) {
  return String(text || '')
    .replace(/\r\n?/g, '\n')
    .replace(/ /g, ' ')
    .replace(/[ \t\f\v]+/g, ' ')
    .split('\n').map(l => l.trim()).join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function decodeEntities(s) {
  const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', frac12: '½', frac14: '¼', frac34: '¾', deg: '°', ndash: '–', mdash: '—', hellip: '…', rsquo: '’', lsquo: '‘', rdquo: '”', ldquo: '“', times: '×' };
  return s
    .replace(/&#(\d+);/g, (m, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&([a-z0-9]+);/gi, (m, name) => named[name] !== undefined ? named[name] : m);
}

// HTML -> text that still looks like a recipe: block elements become line
// breaks, list items get a bullet, chrome (nav/footer/aside/etc.) is dropped.
function htmlToText(html) {
  let h = String(html || '');
  h = h.replace(/<!--[\s\S]*?-->/g, '');
  h = h.replace(/<(script|style|noscript|svg|iframe|template|nav|footer|aside|form|button|select)\b[^>]*>[\s\S]*?<\/\1>/gi, '\n');
  // Prefer the article/main body if the page has one
  const main = h.match(/<(article|main)\b[^>]*>([\s\S]*?)<\/\1>/i);
  if (main && main[2].replace(/<[^>]+>/g, '').trim().length > 800) h = main[2];
  h = h.replace(/<br\s*\/?>/gi, '\n');
  h = h.replace(/<li\b[^>]*>/gi, '\n• ');
  h = h.replace(/<\/(p|div|ul|ol|h[1-6]|tr|section|article|header|blockquote|dd|dt|figcaption|table)>/gi, '\n');
  h = h.replace(/<(h[1-6])\b[^>]*>/gi, '\n\n');
  h = h.replace(/<[^>]+>/g, ' ');
  h = decodeEntities(h);
  return normalizeText(h);
}

// ── URL Import ────────────────────────────────────────────
//
// Apps Script's UrlFetchApp sends its own User-Agent (custom ones are
// ignored) from Google IP ranges, and many recipe publishers refuse it.
// So the page is fetched through a chain of routes, cheapest first, and
// the first one that yields readable recipe text wins:
//   1. direct fetch from Apps Script
//   2. Jina Reader (r.jina.ai) — renders the page and returns its text
//   3. Claude fetching the page itself (web_fetch server tool)
//   4. the Wayback Machine's latest snapshot
// Only when all four fail does the app offer paste-import.

const BROWSER_HEADERS = {
  'User-Agent':      'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
  'Accept-Language': 'en-US,en;q=0.8,pt;q=0.6'
};

function fetchPage(url, headers) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: headers || BROWSER_HEADERS });
  return { status: res.getResponseCode(), body: res.getContentText() };
}

function looksBlocked(text) {
  const lower = String(text || '').slice(0, 3000).toLowerCase();
  return /just a moment|verify you are human|access denied|enable javascript and cookies|captcha|are you a robot|attention required|request blocked|pardon our interruption|403 forbidden/.test(lower);
}

// HTML page -> { text } | { fail: 'BLOCKED' | 'EMPTY' }
function pageToRecipeText(html) {
  const jsonLd = extractJsonLd(html);
  if (jsonLd) return { text: jsonLd };
  const text = htmlToText(html);
  if (looksBlocked(text)) return { fail: 'BLOCKED' };
  if (text.length < 400) return { fail: 'EMPTY' };
  return { text: text };
}

// Interpret one fetched response according to its route.
function routeResult(name, res) {
  if (res.error) return { fail: 'FETCH_FAILED', note: res.error };
  if (res.status >= 400) {
    const blocked = res.status === 403 || res.status === 401 || res.status === 429 || res.status === 503;
    return { fail: (name === 'direct' && blocked) ? 'BLOCKED' : 'HTTP_' + res.status };
  }
  if (name === 'jina') {
    const text = normalizeText(res.body);
    if (looksBlocked(text)) return { fail: 'BLOCKED' };
    if (text.length < 400) return { fail: 'EMPTY' };
    return { text: text };
  }
  if (name === 'wayback') {
    let snap = null;
    try { const j = JSON.parse(res.body); snap = j && j.archived_snapshots && j.archived_snapshots.closest; } catch (e) { return { fail: 'BAD_JSON' }; }
    if (!snap || !snap.available || !snap.url) return { fail: 'NO_SNAPSHOT' };
    let page;
    try { page = fetchPage(snap.url.replace(/\/web\/(\d+)\//, '/web/$1id_/')); }
    catch (e) { return { fail: 'FETCH_FAILED', note: e.message }; }
    if (page.status >= 400) return { fail: 'HTTP_' + page.status };
    const r = pageToRecipeText(page.body);
    return r.text ? { text: r.text } : { fail: r.fail };
  }
  const r = pageToRecipeText(res.body);
  return r.text ? { text: r.text } : { fail: r.fail };
}

// The three plain fetches run concurrently (fetchAll), so the slowest one
// bounds the wait instead of their sum. Results are then preferred in order.
function fetchRoutesInParallel(url) {
  const reqs = [
    { name: 'direct',  url: url, headers: BROWSER_HEADERS },
    { name: 'jina',    url: 'https://r.jina.ai/' + url, headers: { 'Accept': 'text/plain', 'X-Return-Format': 'text' } },
    // Availability API answers in about a second; fetching a snapshot that
    // doesn't exist can hang for the full UrlFetch limit (~100s).
    { name: 'wayback', url: 'https://archive.org/wayback/available?url=' + encodeURIComponent(url), headers: { 'Accept': 'application/json' } }
  ];
  const t0 = Date.now();
  let responses;
  try {
    responses = UrlFetchApp.fetchAll(reqs.map(r => ({ url: r.url, muteHttpExceptions: true, followRedirects: true, headers: r.headers })));
  } catch (e) {
    // fetchAll rejects the whole batch if any URL is malformed; fall back to one-by-one.
    responses = reqs.map(r => { try { return UrlFetchApp.fetch(r.url, { muteHttpExceptions: true, followRedirects: true, headers: r.headers }); } catch (err) { return { error: err.message }; } });
  }
  const ms = Date.now() - t0;
  return reqs.map((r, i) => {
    const res = responses[i];
    let raw;
    if (!res) raw = { error: 'no response' };
    else if (res.error) raw = { error: res.error };
    else { try { raw = { status: res.getResponseCode(), body: res.getContentText() }; } catch (e) { raw = { error: e.message }; } }
    const out = routeResult(r.name, raw);
    out.name = r.name; out.status = raw.status; out.ms = ms;
    return out;
  });
}

function importFromUrl(url) {
  url = String(url || '').trim();
  if (!/^https?:\/\//i.test(url)) {
    return { success: false, code: 'BAD_URL', error: 'That doesn\'t look like a web address. It should start with http:// or https://.' };
  }
  const started = Date.now();
  logEvent('INFO', 'importUrl', 'Start: ' + url);

  const tried = [];
  let sawBlock = false;
  let walled = false;

  // 1–3 in parallel: direct, Jina Reader, Wayback Machine
  const results = fetchRoutesInParallel(url);
  logEvent('INFO', 'importUrl', 'Fetched in ' + results[0].ms + 'ms — ' + results.map(r => r.name + ':' + (r.text ? 'ok(' + r.text.length + ')' : r.fail + (r.status ? '/' + r.status : '') + (r.note ? ' (' + String(r.note).slice(0, 60) + ')' : ''))).join(', '));
  for (let i = 0; i < results.length; i++) {
    const r = results[i];
    if (!r.text) { tried.push(r.name + ':' + r.fail); if (r.fail === 'BLOCKED') sawBlock = true; if (r.status === 402 || r.status === 451) walled = true; continue; }
    const out = callClaudeWithText(r.text.slice(0, TEXT_CHAR_LIMIT), url, 'importUrl');
    if (out.success) { out.via = r.name; logEvent('INFO', 'importUrl', 'OK via ' + r.name + ' in ' + (Date.now() - started) + 'ms: ' + out.recipe.title); return out; }
    tried.push(r.name + ':' + (out.code || 'claude'));
    if (out.code === 'API_AUTH' || out.code === 'CONFIG' || out.code === 'API_BUSY') return out;
  }

  // 4. Let Claude fetch the page itself
  const viaClaude = importViaClaudeFetch(url);
  if (viaClaude.success) { viaClaude.via = 'web_fetch'; logEvent('INFO', 'importUrl', 'OK via web_fetch in ' + (Date.now() - started) + 'ms: ' + viaClaude.recipe.title); return viaClaude; }
  tried.push('web_fetch:' + (viaClaude.code || '?'));
  if (viaClaude.code === 'API_AUTH' || viaClaude.code === 'CONFIG') return viaClaude;

  logEvent('WARN', 'importUrl', 'All routes failed in ' + (Date.now() - started) + 'ms for ' + url + ' — ' + tried.join(', '));
  let host = url; try { host = url.split('/')[2].replace(/^www\./, ''); } catch (e) {}
  return {
    success: false,
    code: (walled || sawBlock) ? 'BLOCKED' : 'UNREACHABLE',
    fallback: 'paste',
    tried: tried,
    error: walled ? host + ' walls off every automated reader.' : sawBlock ? 'This site blocks automated readers, even through the fallbacks.' : 'Couldn\'t get readable recipe text from that page.'
  };
}

// Claude fetches the URL with its own web_fetch tool, then extracts.
function importViaClaudeFetch(url) {
  const content = [{
    type: 'text',
    text: 'Fetch this recipe page with the web_fetch tool, then extract the recipe from it.\nURL: ' + url + '\n\n' + recipePrompt()
  }];
  const result = callClaude(content, 'importUrl:web_fetch', {
    tools: [{ type: 'web_fetch_20260209', name: 'web_fetch', max_uses: 2, max_content_tokens: 40000 }]
  });
  if (result.success) result.recipe.source_url = url;
  return result;
}

// ── JSON-LD extraction ────────────────────────────────────

function isRecipeNode(d) {
  if (!d || typeof d !== 'object') return false;
  const t = d['@type'];
  return t === 'Recipe' || (Array.isArray(t) && t.indexOf('Recipe') !== -1);
}

function findRecipeNode(data, depth) {
  depth = depth || 0;
  if (!data || depth > 4) return null;
  if (Array.isArray(data)) {
    for (let i = 0; i < data.length; i++) {
      const r = findRecipeNode(data[i], depth + 1);
      if (r) return r;
    }
    return null;
  }
  if (typeof data !== 'object') return null;
  if (isRecipeNode(data)) return data;
  if (data['@graph']) return findRecipeNode(data['@graph'], depth + 1);
  if (data.mainEntity) return findRecipeNode(data.mainEntity, depth + 1);
  if (data.mainEntityOfPage && typeof data.mainEntityOfPage === 'object') return findRecipeNode(data.mainEntityOfPage, depth + 1);
  return null;
}

function extractJsonLd(html) {
  const regex = /<script[^>]*type\s*=\s*["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = regex.exec(html)) !== null) {
    try {
      const data = JSON.parse(match[1].trim());
      const recipe = findRecipeNode(data, 0);
      if (recipe) return jsonLdToText(recipe);
    } catch (e) { /* skip malformed JSON-LD */ }
  }
  return null;
}

// ISO 8601 duration ("PT1H30M") -> "1 hr 30 min"
function isoDuration(s) {
  if (!s || typeof s !== 'string') return '';
  const m = s.match(/^P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?/i);
  if (!m) return s;
  const parts = [];
  if (m[1]) parts.push(m[1] + ' day' + (m[1] === '1' ? '' : 's'));
  if (m[2]) parts.push(m[2] + ' hr');
  if (m[3]) parts.push(m[3] + ' min');
  return parts.join(' ');
}

function ldString(v) {
  if (v == null) return '';
  if (typeof v === 'string') return decodeEntities(v.replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim();
  if (Array.isArray(v)) return v.map(ldString).filter(Boolean).join(', ');
  if (typeof v === 'object') return ldString(v.name || v.text || v['@value'] || '');
  return String(v);
}

function ldInstructions(ins, out, depth) {
  depth = depth || 0;
  if (!ins || depth > 3) return;
  if (typeof ins === 'string') {
    ins.split(/\n+/).forEach(l => { const t = ldString(l).trim(); if (t) out.push(t); });
    return;
  }
  if (Array.isArray(ins)) { ins.forEach(i => ldInstructions(i, out, depth + 1)); return; }
  if (typeof ins === 'object') {
    if (ins.name && ins.itemListElement) out.push('## ' + ldString(ins.name));
    if (ins.itemListElement) { ldInstructions(ins.itemListElement, out, depth + 1); return; }
    const t = ldString(ins.text || ins.name);
    if (t) out.push(t);
  }
}

// Render a schema.org Recipe as compact text. Strips reviews, ratings,
// images, video and publisher noise that used to eat the size cap.
function jsonLdToText(r) {
  const lines = [];
  lines.push('Title: ' + ldString(r.name));
  if (r.description)   lines.push('Description: ' + ldString(r.description));
  if (r.recipeCategory) lines.push('Category: ' + ldString(r.recipeCategory));
  if (r.recipeCuisine)  lines.push('Cuisine: ' + ldString(r.recipeCuisine));
  if (r.recipeYield)    lines.push('Yield: ' + ldString(r.recipeYield));
  if (r.prepTime)       lines.push('Prep time: ' + isoDuration(r.prepTime));
  if (r.cookTime)       lines.push('Cook time: ' + isoDuration(r.cookTime));
  if (r.totalTime)      lines.push('Total time: ' + isoDuration(r.totalTime));
  lines.push('');
  lines.push('Ingredients:');
  const ings = r.recipeIngredient || r.ingredients || [];
  (Array.isArray(ings) ? ings : [ings]).forEach(i => { const t = ldString(i).trim(); if (t) lines.push('- ' + t); });
  lines.push('');
  lines.push('Instructions:');
  const steps = [];
  ldInstructions(r.recipeInstructions, steps, 0);
  let n = 0;
  steps.forEach(s => lines.push(s.indexOf('## ') === 0 ? s : (++n) + '. ' + s));
  const notes = [];
  if (r.keywords) notes.push('Keywords: ' + ldString(r.keywords));
  if (r.suitableForDiet) notes.push('Diet: ' + ldString(r.suitableForDiet));
  if (notes.length) { lines.push(''); lines.push(notes.join('\n')); }
  return normalizeText(lines.join('\n'));
}

// ── Text Import ───────────────────────────────────────────

function importFromText(text, url) {
  const clean = normalizeText(text);
  if (clean.length < 40) {
    return { success: false, code: 'EMPTY', error: 'That\'s not enough text to be a recipe. Paste the ingredients and method.' };
  }
  return callClaudeWithText(clean.slice(0, TEXT_CHAR_LIMIT), url, 'importText');
}

// With autosave, an import lands in the book immediately and the response
// carries the new id, so a caller with no review UI (the iOS Shortcut) can
// open the recipe straight away via ?recipe=<id>.
function maybeAutosave(result, autosave) {
  if (!result.success || !isTruthy(autosave)) return result;
  const saved = addRecipe(result.recipe);
  result.id = saved.id;
  result.saved = true;
  result.open_url = 'https://zecipes.andrewgoncalves.com/?recipe=' + saved.id;
  logEvent('INFO', 'autosave', 'Saved via import: ' + result.recipe.title);
  return result;
}

function isTruthy(v) { return v === true || v === 1 || /^(1|true|yes)$/i.test(String(v || '')); }

function callClaudeWithText(text, url, action) {
  const content = [{ type: 'text', text: recipePrompt() + '\n\nText:\n' + text }];
  const result = callClaude(content, action);
  if (result.success) result.recipe.source_url = url || '';
  return result;
}

// ── Image Import ──────────────────────────────────────────

function importFromImages(params) {
  const images = [];

  // New format: images: [{ data, mime }] (JSON body)
  if (Array.isArray(params.images)) {
    params.images.forEach(img => {
      if (img && img.data) images.push({ data: stripDataUrl(img.data), mime: img.mime || 'image/jpeg' });
    });
  }
  // Legacy format: imageCount + image_0/mime_0 …
  const count = parseInt(params.imageCount, 10) || 0;
  for (let i = 0; i < count; i++) {
    if (params['image_' + i]) images.push({ data: stripDataUrl(params['image_' + i]), mime: params['mime_' + i] || 'image/jpeg' });
  }

  if (!images.length) return { success: false, code: 'EMPTY', error: 'No photos were received. Try again.' };

  const content = images.slice(0, 6).map(img => ({
    type: 'image',
    source: { type: 'base64', media_type: img.mime, data: img.data }
  }));
  content.push({ type: 'text', text: imagePrompt() + '\n\nExtract the recipe from the image(s) above.' });

  const result = callClaude(content, 'importImages');
  if (result.success) result.recipe.source_url = '';
  return result;
}

function stripDataUrl(s) {
  s = String(s || '');
  const idx = s.indexOf('base64,');
  return idx === -1 ? s : s.slice(idx + 7);
}

// ── Shared Claude call ────────────────────────────────────

function callClaude(content, action, opts) {
  opts = opts || {};
  const apiKey = getConfig('ANTHROPIC_API_KEY');
  if (!apiKey) return { success: false, code: 'CONFIG', error: 'ANTHROPIC_API_KEY is missing from the config sheet.' };

  const model = getConfig('CLAUDE_MODEL', DEFAULT_MODEL);
  let useStructured = STRUCTURED_OUTPUT_MODELS.test(model);

  let response = null, status = 0, lastErr = '';
  const messages = [{ role: 'user', content: content }];
  for (let attempt = 0; attempt < 4; attempt++) {
    const payload = {
      model:      model,
      max_tokens: 4000,
      messages:   messages
    };
    if (opts.tools) payload.tools = opts.tools;
    if (useStructured) payload.output_config = { format: { type: 'json_schema', schema: recipeSchema() } };

    let apiRes;
    try {
      apiRes = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        headers: {
          'x-api-key':         apiKey,
          'anthropic-version': '2023-06-01',
          'content-type':      'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
    } catch (e) {
      lastErr = e.message;
      logEvent('ERROR', action, 'API call threw: ' + e.message);
      Utilities.sleep(1500);
      continue;
    }

    status = apiRes.getResponseCode();
    try { response = JSON.parse(apiRes.getContentText()); }
    catch (e) { response = null; lastErr = 'Non-JSON API response'; }

    if (status === 200 && response) {
      // Server-side tool loop hit its iteration cap: resend once so it resumes.
      if (response.stop_reason === 'pause_turn' && messages.length === 1) {
        messages.push({ role: 'assistant', content: response.content });
        continue;
      }
      break;
    }

    const msg = response && response.error ? String(response.error.message || '') : ('HTTP ' + status);
    lastErr = msg;
    logEvent('ERROR', action, 'API ' + status + ': ' + msg.slice(0, 300));

    // Model doesn't accept structured output -> retry without it
    if (status === 400 && useStructured && /output_config|output_format|json_schema/i.test(msg)) {
      useStructured = false;
      continue;
    }
    if (status === 429 || status === 529 || status >= 500) {
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }
    break; // other 4xx — don't retry
  }

  if (status !== 200 || !response) {
    if (status === 401) return { success: false, code: 'API_AUTH', error: 'The Anthropic API key was rejected. Check the config sheet.' };
    if (status === 429 || status === 529) return { success: false, code: 'API_BUSY', error: 'The AI service is busy right now. Try again in a minute.' };
    if (status === 400 || status === 404) return { success: false, code: 'API_REQUEST', error: 'The AI request was rejected: ' + lastErr.slice(0, 160) };
    return { success: false, code: 'API_ERROR', error: 'The AI service didn\'t respond (' + (lastErr || 'unknown error').slice(0, 120) + ').' };
  }

  if (response.stop_reason === 'refusal') {
    logEvent('WARN', action, 'Refusal');
    return { success: false, code: 'REFUSED', error: 'The AI declined to read this content.' };
  }
  if (!response.content || !response.content.length) {
    logEvent('ERROR', action, 'Unexpected response: ' + JSON.stringify(response).slice(0, 300));
    return { success: false, code: 'API_ERROR', error: 'Unexpected response from the AI service.' };
  }
  const fetchErr = response.content.find(c => c.type === 'web_fetch_tool_result' && c.content && !Array.isArray(c.content) && c.content.error_code);
  if (fetchErr) {
    logEvent('WARN', action, 'web_fetch error: ' + fetchErr.content.error_code);
    return { success: false, code: 'FETCH_FAILED', error: 'Claude couldn\'t fetch that page (' + fetchErr.content.error_code + ').' };
  }
  const raw = response.content.filter(c => c.type === 'text').map(c => c.text).join('').trim();
  if (response.stop_reason === 'max_tokens') logEvent('WARN', action, 'Output truncated at max_tokens');

  const parsed = parseRecipeJson(raw);
  if (!parsed) {
    logEvent('ERROR', action, 'Could not parse recipe: ' + raw.slice(0, 300));
    return { success: false, code: 'PARSE', error: 'The AI answered but not in a usable format. Try again, or paste less surrounding text.' };
  }

  const recipe = normalizeRecipe(parsed);
  if (!recipe.title && !recipe.ingredients) {
    return { success: false, code: 'NO_RECIPE', error: 'No recipe was found in that content.' };
  }
  logEvent('INFO', action, 'Extracted: ' + recipe.title + ' (' + model + ')');
  return { success: true, recipe, model };
}

function parseRecipeJson(raw) {
  if (!raw) return null;
  let s = raw.replace(/```(?:json)?/gi, '').trim();
  try { return JSON.parse(s); } catch (e) { /* try harder */ }
  const first = s.indexOf('{'), last = s.lastIndexOf('}');
  if (first !== -1 && last > first) {
    try { return JSON.parse(s.slice(first, last + 1)); } catch (e) { /* give up */ }
  }
  return null;
}

function toLines(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return v.map(x => typeof x === 'object' && x !== null ? (x.text || x.name || JSON.stringify(x)) : String(x)).join('\n');
  return String(v);
}

function normalizeRecipe(r) {
  const out = {
    title:        String(r.title || '').trim(),
    category:     String(r.category || '').trim(),
    servings:     String(r.servings || '').trim(),
    prep_time:    String(r.prep_time || '').trim(),
    cook_time:    String(r.cook_time || '').trim(),
    ingredients:  normalizeText(toLines(r.ingredients)),
    instructions: normalizeText(toLines(r.instructions)),
    notes:        normalizeText(toLines(r.notes))
  };
  // Strip step numbering / bullets the model might have added anyway
  out.instructions = out.instructions.split('\n')
    .map(l => l.replace(/^\s*(?:step\s*)?\d+[\.\):]\s*/i, '').replace(/^[•\-*]\s*/, '').trim())
    .filter(Boolean).join('\n');
  out.ingredients = out.ingredients.split('\n')
    .map(l => l.replace(/^[•\-*]\s*/, '').trim())
    .filter(Boolean).join('\n');
  // Snap category to the allowed list (case-insensitive)
  const hit = CATEGORIES.find(c => c.toLowerCase() === out.category.toLowerCase());
  if (hit) out.category = hit;
  return out;
}

// ── doGet — read-only actions ─────────────────────────────

function doGet(e) {
  try {
    const p      = (e && e.parameter) || {};
    const action = p.action;

    if (!action) {
      return ContentService
        .createTextOutput('Zecipes API — use the app at zecipes.andrewgoncalves.com')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    if (action === 'list') {
      return jsonResponse({ success: true, recipes: getAllRecipes() });
    }
    if (action === 'share') {
      const recipe = getRecipeById(p.id);
      if (!recipe) return jsonResponse({ success: false, error: 'Recipe not found' });
      return jsonResponse({ success: true, recipe });
    }
    if (action === 'mealplan') {
      return jsonResponse({ success: true, entries: getMealplan(p.start, p.end) });
    }
    if (action === 'shoppingList') {
      return jsonResponse(getShoppingList(p.start, p.end));
    }
    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    logEvent('ERROR', 'doGet', err.message);
    return jsonResponse({ success: false, error: 'Server error: ' + err.message });
  }
}

// ── doPost — all mutations + large payloads ───────────────

function doPost(e) {
  let p = {};
  try {
    p = parseRequest(e);
    const action = p.action;

    if (action === 'checkPassword') {
      const match = String(p.pw) === String(getConfig('PASSWORD'));
      if (!match) logEvent('WARN', 'auth', 'Failed login attempt');
      return jsonResponse({ success: match });
    }

    if (action === 'add')    return jsonResponse(addRecipe(p));
    if (action === 'update') return jsonResponse(updateRecipe(p));
    if (action === 'delete') return jsonResponse(deleteRecipe(p.id));

    if (action === 'import')           return jsonResponse(maybeAutosave(importFromUrl(p.url), p.autosave));
    if (action === 'importText')       return jsonResponse(maybeAutosave(importFromText(p.text, p.url), p.autosave));
    if (action === 'importFromImages') return jsonResponse(importFromImages(p));

    if (action === 'setMealplan')    return jsonResponse(setMealplan(p.date, p.recipe_id, p.meal_slot || 'dinner'));
    if (action === 'removeMealplan') return jsonResponse(removeMealplan(p.date, p.meal_slot || 'dinner'));

    return jsonResponse({ success: false, error: 'Unknown action' });
  } catch (err) {
    logEvent('ERROR', 'doPost:' + (p.action || '?'), err.message);
    return jsonResponse({ success: false, error: 'Server error: ' + err.message });
  }
}

// ── Response helper ───────────────────────────────────────

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
