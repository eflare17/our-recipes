/**
 * TikTok Recipe Organiser — backend (v2)
 * Paste into a Google Sheet's Apps Script (Extensions > Apps Script), replacing everything.
 *
 * Script properties (Project Settings > Script properties):
 *   GEMINI_KEY   — your Gemini API key (required)
 *
 * Run setup() once after pasting. It creates the sheet, the thumbnail folder
 * and the nightly retry for anything queued.
 */

const SHEET = 'Recipes';
const THUMB_FOLDER = 'Recipe Thumbnails';
const SEARCH_CAP = 4500;          // stops web searching at this many per month (Google gives 5,000 free)
const MODEL_FALLBACK = 'gemini-flash-latest';

const HEAD = ['Added', 'Title', 'Creator', 'TikTok URL', 'Thumbnail', 'Status', 'Tags', 'Time', 'Serves',
  'Ingredients', 'Steps', 'Notes', 'Sources', 'Caption', 'JSON'];
const COL = {}; HEAD.forEach((h, i) => COL[h] = i + 1);

/* ======================= SETUP ======================= */

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  thumbFolder();

  // Nightly retry of queued recipes (3am)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processQueue')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processQueue').timeBased().everyDays(1).atHour(3).create();

  Logger.log('Setup done. Model: ' + pickModel(true) + ' | Searches this month: ' + searchCount());
}

/* ======================= WEB ENDPOINTS ======================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'list') return out({ ok: true, recipes: listRecipes(), searches: searchCount(), cap: SEARCH_CAP });
  if (p.url || p.text) return out(addRecipe(p.url, p.text));
  return out({ ok: true, msg: 'Recipe API running' });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (_) {}
  if (b.action === 'list') return out({ ok: true, recipes: listRecipes() });
  if (b.action === 'update') return out(updateRecipe(b.url, b.changes || {}));
  if (b.action === 'delete') return out(deleteRecipe(b.url));
  if (b.action === 'retry') return out(retryOne(b.url));
  if (b.action === 'fill') return out(fillOne(b.url));
  if (b.action === 'own') return out(addOwn(b.recipe || {}));
  return out(addRecipe(b.url, b.text));
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ======================= ADD A RECIPE ======================= */

function addRecipe(urlIn, textIn) {
  const lock = LockService.getScriptLock();
  try {
    const raw = extractUrl([urlIn, textIn].filter(Boolean).join(' '));
    if (!raw) return { ok: false, error: 'No TikTok link found' };

    const url = resolveUrl(raw);
    const dup = findRow(url);
    if (dup) return { ok: true, duplicate: true, recipe: rowRecipe(dup) };

    const meta = tiktokMeta(url);
    if (!meta.caption && !meta.thumb) return { ok: false, error: "Couldn't read that TikTok — is it public?" };
    meta.thumbDrive = saveThumb(meta.thumb, videoId(url));

    // Write a placeholder row first so nothing is lost if the AI step fails
    lock.waitLock(20000);
    const row = appendRecipe(baseRecipe(url, meta, 'queued'));
    lock.releaseLock();

    return quickPass(row);
  } catch (err) {
    try { lock.releaseLock(); } catch (_) {}
    return { ok: false, error: String((err && err.message) || err) };
  }
}

// Step 1 (fast): read only what the caption says, no web search.
// Returns straight away. If there are gaps the recipe is marked "filling"
// and the app then calls fillOne() to look up 3 similar recipes.
function quickPass(row) {
  const rec = rowRecipe(row);
  try {
    const r = extractRecipe(rec.caption, rec.author, { captionOnly: true });
    const complete = r.status === 'complete';
    Object.assign(rec, {
      title: r.title || rec.title,
      status: complete ? 'complete' : 'filling',
      tags: r.tags || [], time: r.time || '', serves: r.serves || '',
      ingredients: r.ingredients || [], steps: r.steps || [],
      notes: r.notes || '', sources: [], searched: false, error: ''
    });
    writeRecipe(findRow(rec.url) || row, rec);
    return { ok: true, filling: !complete, recipe: rec };
  } catch (err) {
    const msg = String((err && err.message) || err);
    rec.status = 'queued';
    rec.error = /quota|exhausted|429|rate/i.test(msg) ? 'Daily limit hit — will finish tonight' : msg;
    writeRecipe(findRow(rec.url) || row, rec);
    return { ok: true, queued: true, recipe: rec };
  }
}

// Step 2: fill the gaps from 3 similar recipes online
function fillOne(url) {
  const row = findRow(url);
  if (!row) return { ok: false, error: 'Not found' };
  const rec = rowRecipe(row);
  if (rec.status !== 'filling' && rec.status !== 'queued') return { ok: true, recipe: rec };
  return runAi(row);
}

// Full pass with web search on one sheet row. Leaves it "queued" if the daily limit is hit.
function runAi(row) {
  const rec = rowRecipe(row);
  try {
    const r = extractRecipe(rec.caption, rec.author);
    Object.assign(rec, {
      title: r.title || rec.title,
      status: r.status === 'complete' ? 'complete' : 'partial',
      tags: r.tags || [], time: r.time || '', serves: r.serves || '',
      ingredients: r.ingredients || [], steps: r.steps || [],
      notes: r.notes || '', sources: r.sources || [],
      searched: !!r.searched, error: ''
    });
    writeRecipe(findRow(rec.url) || row, rec);
    return { ok: true, recipe: rec };
  } catch (err) {
    const msg = String((err && err.message) || err);
    rec.status = 'queued';
    rec.error = /quota|exhausted|429|rate/i.test(msg) ? 'Daily limit hit — will finish tonight' : msg;
    writeRecipe(findRow(rec.url) || row, rec);
    return { ok: true, queued: true, recipe: rec };
  }
}

// Nightly trigger: retry everything still queued (stops early if limit hit again)
function processQueue() {
  const sh = sheet();
  const n = sh.getLastRow();
  for (let row = 2; row <= n; row++) {
    const st = sh.getRange(row, COL.Status).getValue();
    if (st !== 'queued' && st !== 'filling') continue;
    const res = runAi(row);
    if (res.queued && /limit/i.test(res.recipe.error || '')) break;
    Utilities.sleep(13000); // stay under ~5 requests/minute
  }
}

function retryOne(url) {
  const row = findRow(url);
  if (!row) return { ok: false, error: 'Not found' };
  return runAi(row);
}

/* ======================= TIKTOK ======================= */

function extractUrl(s) {
  const m = String(s || '').match(/https?:\/\/[^\s]*tiktok\.com[^\s]*/i);
  return m ? m[0].replace(/[)\].,]+$/, '') : '';
}

// Turns vt.tiktok.com / vm.tiktok.com short links into the full video link
function resolveUrl(url) {
  let u = url;
  for (let i = 0; i < 5; i++) {
    if (/\/video\/\d+/.test(u)) break;
    const res = UrlFetchApp.fetch(u, { followRedirects: false, muteHttpExceptions: true });
    const h = res.getAllHeaders();
    const loc = h.Location || h.location;
    if (!loc) break;
    u = Array.isArray(loc) ? loc[0] : loc;
  }
  return u.split('?')[0];
}

function videoId(url) {
  const m = String(url).match(/\/video\/(\d+)/);
  return m ? m[1] : '';
}

// Reads the caption, creator and thumbnail. Tries TikTok's embed info first,
// then falls back to the video page itself (TikTok often leaves the embed info blank).
function tiktokMeta(url) {
  const out = { caption: '', author: '', thumb: '' };
  const handle = String(url).match(/tiktok\.com\/@([^/?#]+)/);
  if (handle) out.author = handle[1];

  try {
    const res = UrlFetchApp.fetch('https://www.tiktok.com/oembed?url=' + encodeURIComponent(url), { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const j = JSON.parse(res.getContentText());
      out.caption = j.title || '';
      out.author = j.author_unique_id || out.author || j.author_name || '';
      out.thumb = j.thumbnail_url || '';
    }
  } catch (_) {}

  if (!out.caption || !out.thumb) {
    try {
      const res = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36',
          'Accept-Language': 'en-AU,en;q=0.9'
        }
      });
      const h = res.getContentText();
      const meta = name => {
        const m = h.match(new RegExp('<meta[^>]+(?:property|name)="' + name + '"[^>]+content="([^"]*)"', 'i')) ||
                  h.match(new RegExp('<meta[^>]+content="([^"]*)"[^>]+(?:property|name)="' + name + '"', 'i'));
        return m ? decodeHtml(m[1]) : '';
      };
      const desc = h.match(/"desc":"((?:[^"\\]|\\.)*)"/);
      const fullDesc = desc ? JSON.parse('"' + desc[1] + '"') : '';
      out.caption = out.caption || fullDesc || meta('og:description') || meta('description') || meta('og:title');
      out.thumb = out.thumb || meta('og:image');
    } catch (_) {}
  }
  return out;
}

function decodeHtml(s) {
  return String(s)
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&#x27;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
}

/* ======================= THUMBNAILS → DRIVE ======================= */

function thumbFolder() {
  const it = DriveApp.getFoldersByName(THUMB_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(THUMB_FOLDER);
}

// TikTok thumbnail links expire, so keep our own copy in Drive
function saveThumb(thumbUrl, id) {
  if (!thumbUrl) return '';
  try {
    const res = UrlFetchApp.fetch(thumbUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return '';
    const blob = res.getBlob().setName('tiktok-' + (id || Date.now()) + '.jpg');
    const file = thumbFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
  } catch (_) {
    return '';
  }
}

/* ======================= GEMINI ======================= */

// Picks the newest "gemini-X.Y-flash" model your key can use (checked once a day)
function pickModel(force) {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('MODEL'), when = Number(props.getProperty('MODEL_AT') || 0);
  if (!force && cached && Date.now() - when < 86400000) return cached;
  try {
    const key = props.getProperty('GEMINI_KEY');
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200&key=' + key, { muteHttpExceptions: true });
    const list = (JSON.parse(res.getContentText()).models || [])
      .map(m => m.name.replace('models/', ''))
      .filter(n => /^gemini-\d+(\.\d+)?-flash$/.test(n))
      .sort((a, b) => parseFloat(b.split('-')[1]) - parseFloat(a.split('-')[1]));
    const pick = list[0] || MODEL_FALLBACK;
    props.setProperties({ MODEL: pick, MODEL_AT: String(Date.now()) });
    return pick;
  } catch (_) {
    return cached || MODEL_FALLBACK;
  }
}

function monthKey() { return 'SEARCHES_' + Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyy-MM'); }
function searchCount() { return Number(PropertiesService.getScriptProperties().getProperty(monthKey()) || 0); }
function addSearches(n) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty(monthKey(), String(searchCount() + n));
}

function gemini(prompt, useSearch) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) throw new Error('Add GEMINI_KEY in Project Settings > Script properties');
  const props = PropertiesService.getScriptProperties();
  const noThinkCfg = props.getProperty('NO_THINK_CFG') === '1';
  const call = withThink => {
    const gc = { temperature: 0.3 };
    if (withThink) gc.thinkingConfig = { thinkingLevel: 'low' }; // less "thinking" = faster replies
    const payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: gc };
    if (useSearch) payload.tools = [{ google_search: {} }];
    const r = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + pickModel() + ':generateContent?key=' + key,
      { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
    );
    return { res: r, j: JSON.parse(r.getContentText()) };
  };
  let { res, j } = call(!noThinkCfg);
  if (j.error && !noThinkCfg && /think/i.test(j.error.message || '')) {
    props.setProperty('NO_THINK_CFG', '1'); // this model doesn't take the setting; stop sending it
    ({ res, j } = call(false));
  }
  if (j.error) throw new Error('Gemini (' + (j.error.status || res.getResponseCode()) + '): ' + j.error.message);
  return j;
}

function extractRecipe(caption, author, opts) {
  opts = opts || {};
  const captionOnly = !!opts.captionOnly;
  const canSearch = !captionOnly && searchCount() < SEARCH_CAP;

  const prompt =
`You turn TikTok food videos into clean, cookable recipes.

TikTok caption:
"""${caption}"""
Creator: ${author}

1. Work out exactly what dish this is.
2. Pull every ingredient, amount and step the caption gives. Mark these "caption".
3. ${captionOnly
    ? 'Do NOT fill any gaps and do not invent anything. Only include what the caption actually says.'
    : 'If ingredients, amounts or steps are missing, ' + (canSearch
      ? 'search the web, find 3 good recipes for the same dish, and fill the gaps using what most of them agree on'
      : 'fill the gaps with the most standard version of this dish') + '. Mark these "web". Anything from the caption always beats the filled-in version.'}
4. Use Australian measurements and names (g, ml, cups, °C, capsicum, coriander, spring onion).
5. ${captionOnly
    ? 'status = "complete" only if the caption gives every ingredient with amounts AND the method steps; otherwise "partial".'
    : 'status = "complete" if the caption had the full recipe, "partial" if you filled gaps.'}
6. tags: 2–4 from Breakfast, Lunch, Dinner, Dessert, Snack, Sides, Quick, Healthy, Chicken, Beef, Pork, Seafood, Veggie, Pasta, Air fryer, Slow cooker, Baking.
7. If this isn't a food recipe at all, set title to "Not a recipe" and leave the lists empty.

Reply with ONLY this JSON, no markdown, no extra text:
{"title":"","status":"complete","serves":"","time":"","tags":["Dinner"],"ingredients":[{"text":"","from":"caption"}],"steps":[{"text":"","from":"caption"}],"notes":""}`;

  let j, searched = false;
  if (captionOnly) {
    j = gemini(prompt, false);
  } else if (canSearch) {
    try {
      j = gemini(prompt, true);
      searched = true;
    } catch (err) {
      // Web search isn't available (free tier, or search quota used up):
      // fill gaps without it. Only a limit on this plain call queues the recipe.
      j = gemini(prompt.replace('search the web, find 3 good recipes for the same dish, and fill the gaps using what most of them agree on',
        'fill the gaps with the most standard version of this dish'), false);
    }
  } else {
    j = gemini(prompt, false);
  }

  const c = j.candidates && j.candidates[0];
  if (!c || !c.content) throw new Error('Gemini returned nothing');
  const txt = (c.content.parts || []).map(p => p.text || '').join('');
  const s = txt.indexOf('{'), e = txt.lastIndexOf('}');
  if (s < 0 || e < 0) throw new Error("Gemini reply wasn't JSON");
  const r = JSON.parse(txt.slice(s, e + 1));

  const gm = c.groundingMetadata || {};
  if (searched) addSearches(Math.max(1, (gm.webSearchQueries || []).length));
  const seen = {};
  r.sources = (gm.groundingChunks || []).filter(x => x.web).map(x => ({ title: x.web.title || 'Recipe', url: x.web.uri || '' }))
    .filter(x => !seen[x.title] && (seen[x.title] = true)).slice(0, 3);
  r.searched = searched;
  return r;
}

/* ======================= SHEET ======================= */

function sheet() {
  let sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  if (!sh) { setup(); sh = SpreadsheetApp.getActive().getSheetByName(SHEET); }
  return sh;
}

function baseRecipe(url, meta, status) {
  return {
    added: new Date().toISOString(), title: 'Saving…', author: meta.author, url: url,
    videoId: videoId(url), thumb: meta.thumbDrive || meta.thumb, status: status,
    tags: [], time: '', serves: '', ingredients: [], steps: [], notes: '', sources: [],
    caption: meta.caption, cooked: 0, error: ''
  };
}

function rowValues(rec) {
  const est = x => (x.from === 'web' ? ' (est.)' : '');
  return [
    new Date(rec.added), rec.title, rec.author, rec.url, rec.thumb, rec.status, (rec.tags || []).join(', '),
    rec.time, rec.serves,
    (rec.ingredients || []).map(x => '• ' + x.text + est(x)).join('\n'),
    (rec.steps || []).map((x, i) => (i + 1) + '. ' + x.text + est(x)).join('\n'),
    rec.notes, (rec.sources || []).map(x => x.title + ' — ' + x.url).join('\n'),
    rec.caption, JSON.stringify(rec)
  ];
}

function appendRecipe(rec) {
  const sh = sheet();
  sh.appendRow(rowValues(rec));
  return sh.getLastRow();
}

function writeRecipe(row, rec) {
  sheet().getRange(row, 1, 1, HEAD.length).setValues([rowValues(rec)]);
}

function rowRecipe(row) {
  try { return JSON.parse(sheet().getRange(row, COL.JSON).getValue()); } catch (_) { return null; }
}

function findRow(url) {
  const sh = sheet();
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(2, COL['TikTok URL'], n - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === url) return i + 2;
  return 0;
}

function listRecipes() {
  const sh = sheet();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, COL.JSON, n - 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch (_) { return null; } })
    .filter(Boolean)
    .reverse(); // newest first
}

// Edits from the app (only these fields can change)
const EDITABLE = ['cooked', 'title', 'tags', 'notes', 'time', 'serves', 'ingredients', 'steps'];

function updateRecipe(url, changes) {
  const row = findRow(url);
  if (!row) return { ok: false, error: 'Not found' };
  const rec = rowRecipe(row);
  EDITABLE.forEach(k => { if (k in changes) rec[k] = changes[k]; });
  if (changes.photo) {
    const t = saveDataUrl(changes.photo, 'photo-' + Date.now());
    if (t) rec.thumb = t;
  }
  if (rec.status === 'partial' && !(rec.ingredients || []).concat(rec.steps || []).some(x => x.from === 'web')) {
    rec.status = 'complete'; // all estimates replaced with her own
  }
  rec.edited = new Date().toISOString();
  writeRecipe(row, rec);
  return { ok: true, recipe: rec };
}

// A recipe she writes herself (no TikTok)
function addOwn(r) {
  if (!r.title) return { ok: false, error: 'Give it a name' };
  const id = 'own:' + Date.now();
  const rec = {
    added: new Date().toISOString(), title: r.title, author: '', url: id, videoId: '', own: true,
    thumb: r.photo ? saveDataUrl(r.photo, 'photo-' + Date.now()) : '',
    status: 'complete', tags: r.tags || [], time: r.time || '', serves: r.serves || '',
    ingredients: r.ingredients || [], steps: r.steps || [], notes: r.notes || '',
    sources: [], caption: '', cooked: 0, error: ''
  };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  appendRecipe(rec);
  lock.releaseLock();
  return { ok: true, recipe: rec };
}

// Saves a photo sent from the phone into the thumbnail folder
function saveDataUrl(dataUrl, name) {
  try {
    const m = String(dataUrl).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
    if (!m) return '';
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + '.jpg');
    const file = thumbFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
  } catch (_) {
    return '';
  }
}

function deleteRecipe(url) {
  const row = findRow(url);
  if (!row) return { ok: false, error: 'Not found' };
  sheet().deleteRow(row);
  return { ok: true };
}

/* ======================= TEST ======================= */

// Swap in a real TikTok recipe link, then Run
function testAdd() {
  Logger.log(JSON.stringify(addRecipe('https://www.tiktok.com/@someone/video/1234567890'), null, 2));
}
