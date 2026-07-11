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
    if(action === 'report')    return json_(report_(req));
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
    sh.appendRow(['id','text','lang','mood','ts','hugs','reports']);
  }
  return sh;
}

const REPORT_LIMIT = 3;   // reports needed to auto-hide a message

/* ---------------- actions ---------------- */

function listMessages_(){
  const sh = sheet_();
  const rows = sh.getDataRange().getValues();
  const now = Date.now();
  const out = [];
  let hasExpired = false;

  for(let i = 1; i < rows.length; i++){
    const ts = Number(rows[i][4]);
    const reports = Number(rows[i][6]) || 0;
    if(now - ts <= WEEK_MS){
      if(reports < REPORT_LIMIT){
        out.push({
          id: String(rows[i][0]),
          t:  String(rows[i][1]),
          l:  String(rows[i][2]),
          m:  String(rows[i][3]),
          ts: ts,
          h:  Number(rows[i][5]) || 0
        });
      }
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
    0,
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

function report_(req){
  const id = String(req.id || '');
  if(!id) return {error:'no id'};
  const sh = sheet_();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === id){
      const r = (Number(rows[i][6]) || 0) + 1;
      sh.getRange(i + 1, 7).setValue(r);
      return {ok:true, hidden: r >= REPORT_LIMIT};
    }
  }
  return {error:'not found'};
}

/* ---------------- content guard (server side — the real gate) ----------------
   Layers:
   1) normalize: Arabic-Indic digits → Latin, spelled-out numbers → digits
   2) total digit budget: 7+ digits anywhere = blocked (beats 05-075-673, dots, spaces)
   3) contact info: emails, links, @handles, messenger apps + contact intent
   4) solicitation: selling/promo/dating-contact patterns
------------------------------------------------------------------------------ */

var WORD_DIGITS_ = {
  'zero':'0','one':'1','two':'2','three':'3','four':'4','five':'5',
  'six':'6','seven':'7','eight':'8','nine':'9','oh':'0',
  'صفر':'0','واحد':'1','اثنين':'2','اثنان':'2','ثنين':'2',
  'ثلاثه':'3','ثلاثة':'3','اربعه':'4','اربعة':'4','أربعه':'4','أربعة':'4',
  'خمسه':'5','خمسة':'5','سته':'6','ستة':'6','سبعه':'7','سبعة':'7',
  'ثمانيه':'8','ثمانية':'8','تسعه':'9','تسعة':'9'
};

function normalizeDigits_(text){
  // Arabic-Indic & Extended Arabic-Indic digits → Latin
  var t = text.replace(/[\u0660-\u0669]/g, function(d){ return String(d.charCodeAt(0) - 0x0660); })
              .replace(/[\u06F0-\u06F9]/g, function(d){ return String(d.charCodeAt(0) - 0x06F0); });
  // spelled-out digit words → digits
  t = t.replace(/[A-Za-z\u0600-\u06FF]+/g, function(w){
    var k = w.toLowerCase();
    return WORD_DIGITS_.hasOwnProperty(k) ? WORD_DIGITS_[k] : w;
  });
  return t;
}

var CONTACT_APPS_ = /(whats\s*app|واتس\s*اب|واتساب|واتس|وتساب|telegram|تلي?جرام|تلقرام|تليقرام|snap\s*chat|سناب|snap|insta(gram)?|انستا|انستقرام|انسقرام|tik\s*tok|تيك\s*توك|kik|signal|سيجنال|discord|ديسكورد|imo|ايمو|viber|فايبر|wechat|line\b|face\s*book|فيس\s*بوك|messenger|ماسنجر/i;
var CONTACT_INTENT_ = /(add|follow|dm|inbox|message|msg|text|call|contact|reach|رقمي|رقمى|رقم\s|ضيف|أضيف|اضيف|ضيفو|تواصل|كلمو?ني|راسلو?ني|راسلني|ابعثو|تابعو?ني|عندي\s*(واتس|سناب|انستا))/i;
var SOLICIT_ = /(للبيع|متوفر\s*(توصيل|طلب|كمية)|توصيل\s*لجميع|اسعار\s*خاصه|سعر\s*خاص|dm\s*to\s*(buy|order)|for\s*sale|selling\s|hit\s*me\s*up\s*(for|to)|منتجات\s*خاصه|بضاعه|بضاعة\s*(اصليه|أصلية|متوفره))/i;
var DATING_ = /(تعارف|ابغى\s*(بنت|شب|ولد)|ابي\s*(بنت|شب|ولد)|ودي\s*اتعرف|بنات\s*(واتس|سناب|انستا)|شباب\s*(واتس|سناب)|hook\s*up|sugar\s*(daddy|baby)|onlyfans|اونلي\s*فانز)/i;

function findPII_(text){
  var norm = normalizeDigits_(text);

  // emails (incl. (at)/[at] obfuscation) & links
  if(/[\w.+-]+\s*(@|＠|\(at\)|\[at\])\s*[\w-]+\s*(\.|\(dot\)|\[dot\])\s*\w+/i.test(norm)) return 'email';
  if(/(https?:\/\/|www\.|\.com\b|\.net\b|\.io\b)/i.test(norm)) return 'link';

  // digit budget: 7+ digits total, however they're split or written
  var digits = (norm.match(/\d/g) || []).length;
  if(digits >= 7) return 'phone';

  // @handles
  if(/(^|\s)@[a-z0-9_.]{3,}/i.test(norm)) return 'handle';

  // messenger apps + intent to connect (or any digits at all)
  if(CONTACT_APPS_.test(norm) && (CONTACT_INTENT_.test(norm) || digits >= 3)) return 'contact';

  // "my number is..." style
  if(/(my\s*(number|phone|cell)|رقمي|رقم\s*(جوالي|هاتفي|تلفوني))/i.test(norm)) return 'contact';

  // selling / promo / dating solicitation
  if(SOLICIT_.test(norm)) return 'solicit';
  if(DATING_.test(norm))  return 'solicit';

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
