/**
 * ============================================================
 * PAPER SKIES — Backend (Google Apps Script)
 * ============================================================
 * Storage : Google Sheet ("messages" tab, auto-created)
 * Expiry  : messages older than 7 days are hidden + pruned
 * AI      : translation proxied to Anthropic API (key stays here)
 *
 * SETUP (one time):
 * 1. Create a Google Sheet → Extensions → Apps Script → paste this.
 * 2. Project Settings → Script Properties → add:
 *       ANTHROPIC_API_KEY = sk-ant-xxxxxxxx
 * 3. Deploy → New deployment → Web app
 *       Execute as: Me | Access: Anyone
 * 4. Copy the /exec URL → paste into API_URL in index.html.
 * ============================================================
 */

const WEEK_MS  = 7 * 24 * 60 * 60 * 1000;
const MAX_TEXT = 500;
const MAX_LIST = 200;          // max messages returned to clients
const MODEL    = 'claude-haiku-4-5-20251001';   // fast + cheap for translation

/* ---------------- entry points ---------------- */

function doGet(e){
  return handle_((e && e.parameter) || {});
}

function doPost(e){
  let body = {};
  try{ body = JSON.parse(e.postData.contents); }catch(err){}
  return handle_(body);
}

function handle_(req){
  const action = String(req.action || 'list');
  try{
    if(action === 'list')      return json_(listMessages_());
    if(action === 'add')       return json_(addMessage_(req));
    if(action === 'hug')       return json_(hug_(req));
    if(action === 'translate') return json_(translate_(req));
    return json_({error:'unknown action'});
  }catch(err){
    return json_({error: String(err && err.message || err)});
  }
}

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- sheet helpers ---------------- */

function sheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('messages');
  if(!sh){
    sh = ss.insertSheet('messages');
    sh.appendRow(['id','text','lang','mood','ts','hugs']);
  }
  return sh;
}

/* ---------------- actions ---------------- */

function listMessages_(){
  const sh = sheet_();
  const rows = sh.getDataRange().getValues();
  const now = Date.now();
  const out = [];
  let hasExpired = false;

  for(let i = 1; i < rows.length; i++){
    const ts = Number(rows[i][4]);
    if(now - ts <= WEEK_MS){
      out.push({
        id: String(rows[i][0]),
        t:  String(rows[i][1]),
        l:  String(rows[i][2]),
        m:  String(rows[i][3]),
        ts: ts,
        h:  Number(rows[i][5]) || 0
      });
    }else{
      hasExpired = true;
    }
  }

  // prune expired rows occasionally (cheap housekeeping on read)
  if(hasExpired && rows.length > 50) prune_(sh);

  const total = Number(PropertiesService.getScriptProperties().getProperty('TOTAL') || 0);
  return {ok:true, total: Math.max(total, out.length), messages: out.slice(-MAX_LIST)};
}

function prune_(sh){
  const rows = sh.getDataRange().getValues();
  const now = Date.now();
  for(let i = rows.length - 1; i >= 1; i--){
    if(now - Number(rows[i][4]) > WEEK_MS) sh.deleteRow(i + 1);
  }
}

function addMessage_(req){
  const text = String(req.text || '').trim().slice(0, MAX_TEXT);
  if(text.length < 2) return {error:'empty'};
  if(findPII_(text))  return {error:'pii'};

  // basic flood control: max 300 live messages
  const sh = sheet_();
  if(sh.getLastRow() > 1200) prune_(sh);

  const id = Utilities.getUuid().slice(0, 8);
  const ts = Date.now();
  sh.appendRow([
    id,
    text,
    String(req.lang || 'en').slice(0, 8),
    String(req.mood || '').slice(0, 4),
    ts,
    0
  ]);

  const props = PropertiesService.getScriptProperties();
  props.setProperty('TOTAL', String(Number(props.getProperty('TOTAL') || 0) + 1));

  return {ok:true, id:id, ts:ts};
}

function hug_(req){
  const id = String(req.id || '');
  if(!id) return {error:'no id'};
  const sh = sheet_();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === id){
      const h = (Number(rows[i][5]) || 0) + 1;
      sh.getRange(i + 1, 6).setValue(h);
      return {ok:true, h:h};
    }
  }
  return {error:'not found'};
}

/* ---------------- PII guard (server side) ---------------- */

function findPII_(text){
  if(/[\w.+-]+@[\w-]+\.[\w.-]+/.test(text)) return 'email';
  if(/(https?:\/\/|www\.)\S+/i.test(text))  return 'link';
  const runs = text.match(/[+\d][\d\s\-().]{5,}\d/g) || [];
  for(let i = 0; i < runs.length; i++){
    if((runs[i].match(/\d/g) || []).length >= 8) return 'phone';
  }
  if(/(^|\s)@[a-z0-9_.]{3,}/i.test(text)) return 'handle';
  return null;
}

/* ---------------- translation (FREE via Google Translate built-in) ----------------
   Uses LanguageApp.translate — free, no API key needed.
   If ANTHROPIC_API_KEY is ever added to Script Properties, Claude is used
   instead for more nuanced, emotional translations. Both are cached 6h.  */

const LANG_CODES = {
  'English':'en','Spanish':'es','Arabic':'ar','French':'fr','German':'de',
  'Portuguese':'pt','Turkish':'tr','Hindi':'hi','Indonesian':'id',
  'Chinese':'zh','Japanese':'ja','Korean':'ko','Russian':'ru'
};

function translate_(req){
  const text = String(req.text || '').slice(0, 600);
  const lang = String(req.lang || 'English').slice(0, 24);
  if(!text) return {error:'empty'};

  const cache = CacheService.getScriptCache();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text + '|' + lang);
  const key = 'tr' + Utilities.base64EncodeWebSafe(digest).slice(0, 24);
  const hit = cache.get(key);
  if(hit) return {ok:true, translation:hit, cached:true};

  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  let out = '';

  if(apiKey){
    out = translateWithClaude_(text, lang, apiKey);
  }
  if(!out){
    const code = LANG_CODES[lang] || 'en';
    out = LanguageApp.translate(text, '', code);   // '' = auto-detect source
  }
  if(!out) return {error:'translation failed'};

  cache.put(key, out, 21600);
  return {ok:true, translation: out};
}

function translateWithClaude_(text, lang, apiKey){
  try{
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {'x-api-key': apiKey, 'anthropic-version': '2023-06-01'},
      payload: JSON.stringify({
        model: MODEL,
        max_tokens: 800,
        messages: [{role:'user', content:
          'Translate the following anonymous message into ' + lang +
          '. Preserve its tone and emotion. Return ONLY the translation, nothing else.\n\n' + text}]
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if(data.error) return '';
    return (data.content || []).map(function(b){ return b.text || ''; }).join('\n').trim();
  }catch(err){
    return '';
  }
}
