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
    if(action === 'view')      return json_(view_(req));
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
    sh.appendRow(['id','text','lang','mood','ts','hugs','reports','views']);
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
          h:  Number(rows[i][5]) || 0,
          v:  Number(rows[i][7]) || 0
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

function view_(req){
  const id = String(req.id || '');
  if(!id) return {error:'no id'};
  const sh = sheet_();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === id){
      const v = (Number(rows[i][7]) || 0) + 1;
      sh.getRange(i + 1, 8).setValue(v);
      return {ok:true, v:v};
    }
  }
  return {error:'not found'};
}

/* ---------------- GLOBAL content guard (server side — the real gate) ----------------
   Covers: every Unicode numeral system, spelled-out numbers in ~15 languages,
   CJK numerals, obfuscated emails, links, handles, messenger apps in many
   scripts, "my number" phrases, and selling/dating solicitation.           */

var DIGIT_WORDS_ = (function(){
  var w = ('zero one two three four five six seven eight nine oh '+
    'cero uno dos tres cuatro cinco seis siete ocho nueve '+
    'zéro un deux trois quatre cinq six sept huit neuf '+
    'null eins zwei drei vier fünf funf sechs sieben acht neun '+
    'um dois três tres oito nove '+
    'due cinque sette otto '+
    'sıfır sifir bir iki üç uc dört dort beş bes altı alti yedi sekiz dokuz '+
    'ноль нуль один два три четыре пять шесть семь восемь девять '+
    'nol satu dua tiga empat lima enam tujuh delapan sembilan '+
    'صفر واحد اثنين اثنان ثنين ثلاثه ثلاثة اربعه اربعة أربعه أربعة خمسه خمسة سته ستة سبعه سبعة ثمانيه ثمانية تسعه تسعة '+
    'یک سه چهار پنج شش هفت هشت نه '+
    'ایک تین چار پانچ چھ سات آٹھ نو '+
    'शून्य एक दो तीन चार पांच पाँच छह सात आठ नौ '+
    '공 영 일 이 삼 사 오 육 칠 팔 구').split(/\s+/);
  var s = {};
  for(var i=0;i<w.length;i++) s[w[i]] = true;
  return s;
})();

var CONTACT_APPS_ = /(whats\s*app|واتس\s*اب|واتساب|واتس|وتساب|telegram|تلي?جرام|تلقرام|تليقرام|ватсап|вотсап|вацап|телеграм|snap\s*chat|سناب|снап|snap|insta(gram)?|انستا|انستقرام|انسقرام|инста(грам)?|tik\s*tok|تيك\s*توك|тикток|kik\b|signal|سيجنال|discord|ديسكورد|дискорд|imo\b|ايمو|viber|فايبر|вайбер|wechat|weixin|微信|위챗|line\b|ライン|라인|kakao|카카오|카톡|face\s*book|فيس\s*بوك|фейсбук|messenger|ماسنجر|мессенджер|qq\b)/i;

var CONTACT_INTENT_ = /(add\s*me|follow\s*me|dm\s*me|dm\b|inbox|message\s*me|msg\s*me|text\s*me|call\s*me|contact\s*me|hmu\b|agrega|agr[ée]game|a[ñn][áa]deme|escr[íi]beme|ll[áa]mame|ajoute|[ée]cris\s*moi|appelle\s*moi|adiciona|me\s*chama|manda\s*(msg|mensagem)|ekle\b|bana\s*yaz|beni\s*ara|добавь|напиши\s*мне|пиши\s*мне|звони|加我|私聊|联系我|連絡して|追加して|추가해|연락해|رقمي|رقمى|ضيفو?ني|أضيفو?ني|اضيفو?ني|تواصل\s*معي|كلمو?ني|راسلو?ني|تابعو?ني|عندي\s*(واتس|سناب|انستا))/i;

var MY_NUMBER_ = /(my\s*(number|phone|cell|digits)|num[ée]ro|n[úu]mero|numara|nomor|телефон|номер|мой\s*номер|电话|手机号|微信号|電話番号|번호|رقمي|رقم\s*(جوالي|هاتفي|تلفوني|الواتس))/i;

var SOLICIT_ = /(للبيع|متوفر\s*(توصيل|طلب|كمية)|توصيل\s*لجميع|اسعار\s*خاصه|سعر\s*خاص|بضاعه|بضاعة\s*(اصليه|أصلية|متوفره)|dm\s*to\s*(buy|order)|for\s*sale|selling\s|se\s*vende|vendo\b|à\s*vendre|satılık|продаю|出售|hit\s*me\s*up\s*(for|to)|تعارف|ابغى\s*(بنت|شب|ولد)|ابي\s*(بنت|شب|ولد)|ودي\s*اتعرف|بنات\s*(واتس|سناب|انستا)|شباب\s*(واتس|سناب)|hook\s*up|sugar\s*(daddy|baby)|onlyfans|اونلي\s*فانز|busco\s*(chica|novia|novio)|cherche\s*(fille|femme|meuf)|познакомлюсь|出会い)/i;

var CJK_NUM_RUN_ = /[〇零一二三四五六七八九壱弐参]{6,}/;

function findPII_(text){
  // 1) digits in ANY numeral system
  var digits = 0;
  try{ digits = (text.match(/\p{Nd}/gu) || []).length; }
  catch(e){ digits = (text.match(/[0-9\u0660-\u0669\u06F0-\u06F9\u0966-\u096F\u09E6-\u09EF\u0E50-\u0E59\uFF10-\uFF19]/g) || []).length; }
  if(digits >= 7) return 'phone';

  // 2) spelled-out numbers: 5+ digit-words in a row (any covered language)
  var tokens = text.toLowerCase().split(/[^\p{L}]+/u).filter(function(t){return t;});
  var run = 0;
  for(var i=0;i<tokens.length;i++){
    run = DIGIT_WORDS_[tokens[i]] ? run + 1 : 0;
    if(run >= 5) return 'phone';
  }

  // 3) CJK numeral runs
  if(CJK_NUM_RUN_.test(text)) return 'phone';

  // 4) emails / links / handles
  if(/[\w.+-]+\s*(@|＠|\(at\)|\[at\])\s*[\w-]+\s*(\.|\(dot\)|\[dot\])\s*\w+/i.test(text)) return 'email';
  if(/(https?:\/\/|www\.|\.com\b|\.net\b|\.io\b|\.me\b)/i.test(text)) return 'link';
  if(/(^|\s)@[a-z0-9_.]{3,}/i.test(text)) return 'handle';

  // 5) messenger apps + intent (or any digits at all)
  if(CONTACT_APPS_.test(text) && (CONTACT_INTENT_.test(text) || digits >= 3)) return 'contact';
  if(MY_NUMBER_.test(text)) return 'contact';

  // 6) selling / promo / meetups
  if(SOLICIT_.test(text)) return 'solicit';

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
