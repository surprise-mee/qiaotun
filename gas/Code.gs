/**
 * 翹臀冒險 — Google Apps Script 後台
 * ──────────────────────────────────────────────
 * 前台（GitHub Pages）呼叫 doPost：註冊／登入、雲端存檔、排行榜、拍手、
 *   讀取動作圖設定、用 Gemini 生成寵物（文字＋每個階段一張圖）。
 * 管理後台：部署網址後面加 ?page=admin
 *
 * 「專案設定 → 指令碼屬性」需要：
 *   ADMIN_PASS  管理後台密碼（自己設）
 *   GEMINI_KEY  Google AI Studio 的 API 金鑰（https://aistudio.google.com/apikey）
 * 其他設定（每月生圖上限、模型、前台網址）都在管理後台「設定」裡改。
 */

const SH = { users: 'users', data: 'data', claps: 'claps', pets: 'pets', usage: 'usage', config: 'config' };
const HEAD = {
  users:  ['名稱', '密碼雜湊', 'token', '建立', '最後上線', '鹽'],
  data:   ['名稱', '存檔', '排行榜摘要', '更新'],
  claps:  ['誰', '拍給誰', '日期'],
  pets:   ['key', '主人', '本機ID', '屬性', '顏色', '突變', '代數', '父母key', '文字', '圖', '狀態', '建立', '更新'],
  usage:  ['月份', '生圖張數'],
  config: ['設定', '值']
};
const STAGES = ['baby', 'teen', 'adult', 'old'];
const DEF = { QUOTA_MONTH: '100', TEXT_MODEL: 'gemini-2.5-flash', IMG_MODEL: 'gemini-2.5-flash-image', SITE_URL: '' };
const TZ = 'Asia/Taipei';

// ================= 入口 =================
function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin').setTitle('翹臀冒險・管理後台')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return out_({ ok: true, app: 'qiaotun', time: new Date().toISOString() });
}

function doPost(e) {
  let r;
  try { r = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, err: '格式錯誤' }); }
  try {
    switch (r.a) {
      case 'register': return out_(locked_(() => register_(r)));
      case 'login':    return out_(locked_(() => login_(r)));
      case 'config':   return out_(publicConfig_());
      case 'save':     return out_(auth_(r, save_));
      case 'board':    return out_(auth_(r, board_));
      case 'clap':     return out_(auth_(r, clap_));
      case 'petgen':   return out_(auth_(r, petgen_));
      case 'admin':    return out_(adminCall_(r));
      default:         return out_({ ok: false, err: '未知的動作' });
    }
  } catch (err) {
    return out_({ ok: false, err: String(err.message || err) });
  }
}

// ================= 帳號 =================
function register_(r) {
  const name = clean_(r.name), pin = String(r.pin || '');
  if (!name) return { ok: false, err: '請輸入冒險者名稱' };
  if (!/^\d{4}$/.test(pin)) return { ok: false, err: '密碼要 4 位數' };
  const sh = sheet_(SH.users);
  if (findRow_(sh, name)) return { ok: false, err: '這個名稱已經有人用了，換一個吧' };
  const salt = Utilities.getUuid(), token = Utilities.getUuid();
  sh.appendRow(["'" + name, hash_(pin, salt), token, new Date(), new Date(), salt]);
  return { ok: true, token: token, name: name, data: null };
}

function login_(r) {
  const name = clean_(r.name), pin = String(r.pin || '');
  const sh = sheet_(SH.users), row = findRow_(sh, name);
  if (!row) return { ok: false, err: '找不到這個冒險者，要先註冊喔' };
  const v = sh.getRange(row, 1, 1, 6).getValues()[0];
  if (hash_(pin, v[5]) !== v[1]) return { ok: false, err: '密碼不對' };
  const token = Utilities.getUuid();
  sh.getRange(row, 3, 1, 3).setValues([[token, v[3], new Date()]]);
  const d = sheet_(SH.data), dr = findRow_(d, name);
  return { ok: true, token: token, name: name, data: dr ? JSON.parse(d.getRange(dr, 2).getValue() || 'null') : null };
}

function auth_(r, fn) {
  const name = clean_(r.name), sh = sheet_(SH.users), row = findRow_(sh, name);
  if (!row || sh.getRange(row, 3).getValue() !== r.token) return { ok: false, err: 'auth', msg: '登入已過期，請重新登入' };
  sh.getRange(row, 5).setValue(new Date());
  return fn(name, r);
}

// ================= 存檔／排行榜 =================
function save_(name, r) {
  const data = JSON.stringify(r.data || {}), sum = JSON.stringify(r.summary || {});
  if (data.length > 48000) return { ok: false, err: '存檔太大' };
  return locked_(() => {
    const sh = sheet_(SH.data), row = findRow_(sh, name);
    if (row) sh.getRange(row, 2, 1, 3).setValues([[data, sum, new Date()]]);
    else sh.appendRow(["'" + name, data, sum, new Date()]);
    return { ok: true };
  });
}

function board_(name) {
  const users = [];
  rows_(SH.data).forEach(v => { if (v[0] && v[2]) { try { const s = JSON.parse(v[2]); s.n = String(v[0]); users.push(s); } catch (e) {} } });
  const today = day_(new Date()), got = {}, mine = [];
  rows_(SH.claps).forEach(c => {
    if (c[0] && day_(c[2]) === today) { got[c[1]] = (got[c[1]] || 0) + 1; if (String(c[0]) === name) mine.push(String(c[1])); }
  });
  users.forEach(u => u.claps = got[u.n] || 0);
  // 自己的寵物最新狀態（管理員重生或下架後，前台會同步）
  const pets = {};
  rows_(SH.pets).forEach(v => { if (String(v[1]) === name) pets[v[2]] = { key: v[0], img: urls_(json_(v[9], {})), status: v[10] }; });
  return { ok: true, users: users, clapped: mine, pets: pets };
}

function clap_(name, r) {
  const to = clean_(r.to);
  if (!to || to === name) return { ok: false, err: '不能拍自己' };
  const today = day_(new Date());
  return locked_(() => {
    const sh = sheet_(SH.claps);
    if (rows_(SH.claps).some(c => String(c[0]) === name && String(c[1]) === to && day_(c[2]) === today)) return { ok: false, err: '今天已經拍過了' };
    sh.appendRow(["'" + name, "'" + to, "'" + today]);
    return { ok: true };
  });
}

// ================= 前台設定（動作圖覆寫） =================
function publicConfig_() {
  return { ok: true, ex: json_(cfg_('EX'), {}), img: json_(cfg_('IMG'), {}), sz: json_(cfg_('IMG_SZ'), {}), v: (cfg_('EX_V') || '0') + '.' + (cfg_('IMG_V') || '0') };
}

// ================= Gemini 寵物 =================
const ATTR_INFO = {
  red:    { n: '力量系', color: 'warm coral red',     egg: 'coral red egg with small dumbbell pattern' },
  mint:   { n: '柔軟系', color: 'soft mint green',    egg: 'mint green egg with wave pattern' },
  butter: { n: '耐力系', color: 'buttery pale yellow', egg: 'pale yellow egg with lightning-bolt pattern' },
  peach:  { n: '翹臀系', color: 'peach pink',         egg: 'peach pink egg with little peach pattern' }
};
const STYLES = ['帥帥可愛', '憨憨可愛', '賤賤可愛', '軟萌可愛', '醜萌可愛', '傲嬌可愛', '呆萌可愛', '酷酷可愛', '迷糊可愛', '浮誇可愛'];

function petgen_(name, r) {
  const p = r.pet || {};
  if (!p.id || !ATTR_INFO[p.attr]) return { ok: false, err: '寵物資料不完整' };
  const key = name + '|' + p.id;
  let rec = getPet_(key);
  if (rec && rec.status === 'hidden') return { ok: true, key: key, t: rec.t, img: {}, hidden: true };
  if (!rec) {
    const parent = p.parentKey ? getPet_(p.parentKey) : null;
    const t = genText_(p, parent);
    rec = { key: key, owner: name, lid: p.id, attr: p.attr, col: p.col || p.attr, mut: p.mut ? 1 : 0, gen: p.gen || 1,
            parent: parent ? parent.key : '', t: t, img: {}, status: 'ok', created: new Date() };
    savePet_(rec);
  }
  const want = (r.stages || []).filter(k => STAGES.indexOf(k) >= 0 && !rec.img[k]);
  let quota = false, err = '';
  for (const k of want) {
    if (!quotaTake_()) { quota = true; break; }
    try {
      genStageImage_(rec, k);
      savePet_(rec);
    } catch (e) {
      quotaGive_(); err = String(e.message || e); break;
    }
  }
  return { ok: true, key: key, t: rec.t, img: urls_(rec.img), quota: quota, err: err };
}

function genText_(p, parent) {
  const used = rows_(SH.pets).map(v => json_(v[8], {}).proto).filter(Boolean).slice(-40);
  const style = STYLES[Math.floor(Math.random() * STYLES.length)];
  const a = ATTR_INFO[p.attr], col = ATTR_INFO[p.col || p.attr];
  const fam = parent ? `這隻是「${parent.t.name}」（${parent.t.proto}）的孩子，第 ${p.gen || 2} 代。物種必須和父母一樣是「${parent.t.proto}」，要保留父母的招牌特徵「${parent.t.mark}」，但牠是全新的個體，名字、個性、故事都要不同。${p.mut ? '牠是突變個體，毛色變成「' + col.color + '」。' : ''}`
    : `請挑一種動物當原型（可以是常見或冷門動物，不要用以下已經出現過的：${used.join('、') || '無'}）。`;
  const prompt = `你是可愛健身手遊《翹臀冒險》的寵物設計師。玩家每次運動會累積能量讓蛋孵化、長大。
請設計一隻新寵物，用繁體中文（台灣用語），全部內容要有趣、溫暖、跟健身／練臀有關。
- 屬性：${a.n}（主色：${col.color}${p.mut ? '，突變色' : ''}）
- 可愛風格：${style}
- ${fam}
- 名字要是「諧音梗」或雙關（跟健身、臀部、動物有關），2–4 個字，好念好記。
- 牠會從幼年 → 青年 → 壯年 → 老年，每個階段外觀要差很多，但一眼看得出是同一隻。
只輸出 JSON，格式如下（不要多餘文字）：
{
 "name": "名字",
 "pun": "名字的諧音梗說明，一句話",
 "style": "${style}",
 "proto": "原型動物（中文）",
 "mark": "全身唯一、各階段都保留的招牌特徵，例如某個斑紋或配件",
 "evo": {"teen": "進化成青年時的變化描述，一句話", "adult": "進化成壯年時的變化描述，一句話"},
 "pers": "個性描述，2–3 句",
 "catch": "口頭禪，一句",
 "bg": "背景故事，3–4 句",
 "story": ["12 段連續的小劇情，每段 1–2 句，從剛孵化到決定陪玩家冒險"],
 "lines": ["10 句陪練時會說的台詞，符合個性"],
 "old": "變成老年時的感人告別：捧出一顆傳承蛋交給玩家，3–4 句",
 "act": {"hi5": ["擊掌時的 3 句反應"], "pat": ["被摸頭的 3 句反應"], "snack": ["吃點心的 3 句反應"]},
 "look": "英文：這隻角色的外觀設定，給畫師用。寫出物種、主色（${col.color}）、眼睛、招牌特徵、體型，約 40 字",
 "stageLook": {"baby": "英文，幼年外觀與動作", "teen": "英文，青年外觀、運動配件與動作", "adult": "英文，壯年外觀、招牌配件與自信姿勢", "old": "英文，老年外觀（白鬍子或白眉、眼鏡等）"}
}`;
  const model = cfg_('TEXT_MODEL') || DEF.TEXT_MODEL;
  const gc = { responseMimeType: 'application/json', temperature: 1.1 };
  // 少想一點、快一點：寫寵物設定不需要深度推理
  if (/^gemini-[3-9]/.test(model)) gc.thinkingConfig = { thinkingLevel: 'low' };
  else if (/^gemini-2\.5-flash/.test(model)) gc.thinkingConfig = { thinkingBudget: 0 };
  let res;
  try { res = gemini_(model, { contents: [{ parts: [{ text: prompt }] }], generationConfig: gc }); }
  catch (e) {
    if (!gc.thinkingConfig || !/400/.test(String(e.message))) throw e;
    delete gc.thinkingConfig;
    res = gemini_(model, { contents: [{ parts: [{ text: prompt }] }], generationConfig: gc });
  }
  const txt = (((res.candidates || [])[0] || {}).content || {}).parts.map(x => x.text || '').join('');
  const t = JSON.parse(txt.replace(/^```json|```$/g, '').trim());
  ['story', 'lines'].forEach(k => { if (!Array.isArray(t[k])) t[k] = []; });
  while (t.story.length < 12) t.story.push(t.lines[t.story.length % Math.max(1, t.lines.length)] || '牠開心地繞著你轉圈。');
  t.style = t.style || style;
  return t;
}

const ART = 'Cute chibi mascot character for a mobile fitness game. Soft pastel colors, hand-drawn look with warm dark-brown outlines, ' +
  'gentle cel shading and soft highlights, rosy blush cheeks, big glossy eyes, a few small doodles around it (tiny hearts, stars, sweat drops). ' +
  'Plain warm cream background (#F8F0E3). One character only, full body, centered, square composition. No text, no letters, no watermark, no frame.';
const STAGE_TXT = {
  baby:  'Life stage: BABY. Tiny, round, newborn and wobbly, sitting down. It wears a broken piece of its eggshell on its head like a little hat (eggshell: %EGG%).',
  teen:  'Life stage: TEEN. Noticeably bigger and more energetic than the baby, new body proportions, wearing a sporty accessory such as a headband or wristbands, doing a cute workout pose (squat, stretch or punch).',
  adult: 'Life stage: ADULT. Strong, confident and fit but still cute, rounder and sturdier body, signature hero accessory (for example a little cape or medal), proud victory pose.',
  old:   'Life stage: ELDER. Old and kind: white eyebrows or white beard, small round glasses, gentle smile, slightly hunched, holding a softly glowing egg in its paws/hands with both arms.'
};

function genStageImage_(rec, k) {
  const t = rec.t, egg = ATTR_INFO[rec.attr].egg;
  let ref = null, refNote = '';
  const i = STAGES.indexOf(k);
  for (let j = i - 1; j >= 0; j--) { if (rec.img[STAGES[j]]) { ref = rec.img[STAGES[j]]; break; } }
  if (ref) refNote = 'The attached image is THIS SAME character at an earlier life stage. Keep the same species, color palette, eye style and signature mark, but redesign the body proportions, pose and outfit so this stage looks clearly different and more grown-up.';
  else if (k === 'baby' && rec.parent) {
    const par = getPet_(rec.parent);
    const pr = par && (par.img.adult || par.img.teen || par.img.baby);
    if (pr) { ref = pr; refNote = 'The attached image is this character\'s PARENT. Draw its newborn child: same species and signature mark, family resemblance, but a new individual.'; }
  }
  const prompt = [ART, 'Character: ' + (t.look || t.proto), (t.stageLook && t.stageLook[k]) || '', STAGE_TXT[k].replace('%EGG%', egg),
    'Signature mark that must be visible: ' + (t.mark || ''), refNote].join('\n');
  const parts = [{ text: prompt }];
  if (ref) {
    const b = DriveApp.getFileById(ref.id).getBlob();
    parts.push({ inlineData: { mimeType: b.getContentType() || 'image/png', data: Utilities.base64Encode(b.getBytes()) } });
  }
  const res = gemini_(cfg_('IMG_MODEL') || DEF.IMG_MODEL, {
    contents: [{ parts: parts }],
    generationConfig: { responseModalities: ['IMAGE'], imageConfig: { aspectRatio: '1:1' } }
  });
  const ps = (((res.candidates || [])[0] || {}).content || {}).parts || [];
  const im = ps.find(x => x.inlineData && x.inlineData.data);
  if (!im) throw new Error('Gemini 沒有回傳圖片' + (res.promptFeedback ? '：' + JSON.stringify(res.promptFeedback) : ''));
  const blob = Utilities.newBlob(Utilities.base64Decode(im.inlineData.data), im.inlineData.mimeType || 'image/png',
    rec.key.replace(/[^\w\-]/g, '_') + '_' + k + '_' + Date.now() + '.png');
  const f = folder_('pets').createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (rec.img[k]) { try { DriveApp.getFileById(rec.img[k].id).setTrashed(true); } catch (e) {} }
  rec.img[k] = { id: f.getId() };
}

function gemini_(model, body) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) throw new Error('還沒設定 GEMINI_KEY');
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key }
  });
  const code = res.getResponseCode(), j = JSON.parse(res.getContentText() || '{}');
  if (code !== 200) throw new Error('Gemini ' + code + '：' + ((j.error && j.error.message) || res.getContentText().slice(0, 200)));
  return j;
}

// ---- 每月生圖上限 ----
function quotaTake_() {
  return locked_(() => {
    const m = Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), sh = sheet_(SH.usage), row = findRow_(sh, m);
    const used = row ? Number(sh.getRange(row, 2).getValue()) : 0, cap = Number(cfg_('QUOTA_MONTH') || DEF.QUOTA_MONTH);
    if (used >= cap) return false;
    if (row) sh.getRange(row, 2).setValue(used + 1); else sh.appendRow(["'" + m, 1]);
    return true;
  });
}
function quotaGive_() {
  locked_(() => {
    const m = Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), sh = sheet_(SH.usage), row = findRow_(sh, m);
    if (row) sh.getRange(row, 2).setValue(Math.max(0, Number(sh.getRange(row, 2).getValue()) - 1));
  });
}
function quotaUsed_() {
  const m = Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), row = rows_(SH.usage).find(v => String(v[0]) === m);
  return row ? Number(row[1]) : 0;
}

// ---- 寵物存取 ----
function getPet_(key) {
  const v = rows_(SH.pets).find(x => String(x[0]) === key);
  if (!v) return null;
  return { key: String(v[0]), owner: String(v[1]), lid: String(v[2]), attr: v[3], col: v[4], mut: v[5], gen: v[6], parent: String(v[7] || ''),
           t: json_(v[8], {}), img: json_(v[9], {}), status: v[10] || 'ok', created: v[11] };
}
function savePet_(rec) {
  locked_(() => {
    const sh = sheet_(SH.pets), row = findRow_(sh, rec.key);
    const vals = ["'" + rec.key, "'" + rec.owner, "'" + rec.lid, rec.attr, rec.col, rec.mut, rec.gen, rec.parent, JSON.stringify(rec.t), JSON.stringify(rec.img), rec.status || 'ok', rec.created || new Date(), new Date()];
    if (row) sh.getRange(row, 1, 1, vals.length).setValues([vals]); else sh.appendRow(vals);
  });
}
function urls_(img) {
  const o = {};
  Object.keys(img || {}).forEach(k => { if (img[k] && img[k].id) o[k] = 'https://lh3.googleusercontent.com/d/' + img[k].id + '=w480'; });
  return o;
}

// ================= 管理後台（Admin.html 用 fetch 呼叫 doPost，a:'admin'） =================
const ADMIN_FNS = { adminLogin: adminLogin, adminOverview: adminOverview, adminPets: adminPets, adminPetRegen: adminPetRegen, adminPetStatus: adminPetStatus,
  adminExGet: adminExGet, adminExUpload: adminExUpload, adminExSet: adminExSet, adminExReset: adminExReset,
  adminImgGet: adminImgGet, adminImgUpload: adminImgUpload, adminImgReset: adminImgReset, adminImgSize: adminImgSize,
  adminSettings: adminSettings, adminSaveSettings: adminSaveSettings, adminTestKey: adminTestKey, adminResetPin: adminResetPin, adminDelete: adminDelete };
function adminCall_(r) {
  const f = ADMIN_FNS[r.fn];
  if (!f) return { ok: false, err: '未知的管理功能' };
  try { return { ok: true, v: f.apply(null, r.args || []) }; } catch (e) { return { ok: false, err: String(e.message || e) }; }
}
function adminLogin(pass) {
  const c = CacheService.getScriptCache(), fails = Number(c.get('adm_fail') || 0);
  if (fails >= 5) throw new Error('密碼錯太多次，請 10 分鐘後再試');
  const p = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS');
  if (!p) throw new Error('還沒在指令碼屬性設定 ADMIN_PASS');
  if (pass !== p) { c.put('adm_fail', String(fails + 1), 600); throw new Error('密碼不對'); }
  c.remove('adm_fail');
  const tok = Utilities.getUuid();
  c.put('adm_' + tok, '1', 21600);
  return tok;
}
function adm_(tok) { if (!tok || !CacheService.getScriptCache().get('adm_' + tok)) throw new Error('請重新登入管理後台'); }

function adminOverview(tok) {
  adm_(tok);
  const now = new Date(), wk0 = weekStart_(now);
  const weeks = []; for (let i = 7; i >= 0; i--) { const d = new Date(wk0); d.setDate(d.getDate() - i * 7); weeks.push(day_(d)); }
  const trend = weeks.map(w => ({ w: w, users: 0, cnt: 0 }));
  const data = {}; rows_(SH.data).forEach(v => { if (v[0]) data[String(v[0])] = json_(v[2], {}); });
  const petCount = {}; rows_(SH.pets).forEach(v => { petCount[v[1]] = (petCount[v[1]] || 0) + 1; });
  const list = rows_(SH.users).filter(v => v[0]).map(v => {
    const n = String(v[0]), s = data[n] || {}, h = s.h || [];
    const seenW = {};
    h.forEach(x => { for (let i = trend.length - 1; i >= 0; i--) { if (x[0] >= trend[i].w) { trend[i].cnt++; seenW[i] = 1; break; } } });
    Object.keys(seenW).forEach(i => trend[i].users++);
    return { name: n, created: v[3] ? Utilities.formatDate(new Date(v[3]), TZ, 'yyyy/MM/dd') : '', seen: v[4] ? Utilities.formatDate(new Date(v[4]), TZ, 'MM/dd HH:mm') : '',
      title: s.t || '', total: h.length, week: h.filter(x => x[0] >= weeks[7]).length, xp: h.reduce((a, x) => a + (x[1] || 0), 0),
      last: h.length ? h[h.length - 1][0] : '', pets: petCount[n] || 0, pet: s.pet || null };
  });
  return { list: list, trend: trend, used: quotaUsed_(), quota: Number(cfg_('QUOTA_MONTH') || DEF.QUOTA_MONTH), sheet: SpreadsheetApp.getActiveSpreadsheet().getUrl() };
}

function adminPets(tok) {
  adm_(tok);
  return rows_(SH.pets).filter(v => v[0]).map(v => {
    const t = json_(v[8], {});
    return { key: String(v[0]), owner: String(v[1]), attr: v[3], col: v[4], mut: v[5], gen: v[6], name: t.name || '', proto: t.proto || '', style: t.style || '',
      pun: t.pun || '', pers: t.pers || '', img: urls_(json_(v[9], {})), status: v[10] || 'ok', created: v[11] ? Utilities.formatDate(new Date(v[11]), TZ, 'yyyy/MM/dd') : '' };
  }).reverse();
}
function adminPetRegen(tok, key, stage) {
  adm_(tok);
  const rec = getPet_(key); if (!rec) throw new Error('找不到寵物');
  if (STAGES.indexOf(stage) < 0) throw new Error('階段不對');
  if (!quotaTake_()) throw new Error('本月生圖額度已用完，請到「設定」調高上限');
  try { genStageImage_(rec, stage); savePet_(rec); } catch (e) { quotaGive_(); throw e; }
  return urls_(rec.img)[stage];
}
function adminPetStatus(tok, key, status) {
  adm_(tok);
  const rec = getPet_(key); if (!rec) throw new Error('找不到寵物');
  rec.status = status === 'hidden' ? 'hidden' : 'ok'; savePet_(rec);
  return rec.status;
}

function adminExGet(tok) { adm_(tok); return { ex: json_(cfg_('EX'), {}), site: cfg_('SITE_URL') || '' }; }
function adminExUpload(tok, id, big, thumb, mime) {
  adm_(tok);
  const f = folder_('exercises');
  const mk = (b64, suf) => { const x = f.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg', id + suf + '_' + Date.now())); x.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return x.getId(); };
  const ex = json_(cfg_('EX'), {}), o = ex[id] || {};
  [o.ub, o.tb].forEach(fid => { if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} } });
  o.ub = mk(big, ''); o.tb = mk(thumb, '_t');
  o.u = 'https://lh3.googleusercontent.com/d/' + o.ub; o.t = 'https://lh3.googleusercontent.com/d/' + o.tb;
  ex[id] = o; setEx_(ex);
  return o;
}
function adminExSet(tok, id, hidden) {
  adm_(tok);
  const ex = json_(cfg_('EX'), {}), o = ex[id] || {};
  o.h = hidden; ex[id] = o; setEx_(ex);
  return o;
}
function adminExReset(tok, id) {
  adm_(tok);
  const ex = json_(cfg_('EX'), {}), o = ex[id];
  if (o) { [o.ub, o.tb].forEach(fid => { if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} } }); delete o.u; delete o.t; delete o.ub; delete o.tb; if (o.h === undefined) delete ex[id]; }
  setEx_(ex);
  return ex[id] || null;
}
// ---- 固定素材（封面、教練、蛋、預設寵物、伸展圖…）----
function adminImgGet(tok) { adm_(tok); return { img: json_(cfg_('IMG'), {}), sz: json_(cfg_('IMG_SZ'), {}), files: json_(cfg_('IMG_F'), {}), site: cfg_('SITE_URL') || '' }; }
function adminImgUpload(tok, key, b64, mime) {
  adm_(tok);
  if (!/^[a-z0-9_]+$/i.test(key)) throw new Error('素材名稱不對');
  const img = json_(cfg_('IMG'), {}), files = json_(cfg_('IMG_F'), {});
  if (files[key]) { try { DriveApp.getFileById(files[key]).setTrashed(true); } catch (e) {} }
  const f = folder_('assets').createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/png', key + '_' + Date.now()));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  files[key] = f.getId(); img[key] = 'https://lh3.googleusercontent.com/d/' + f.getId();
  setCfg_('IMG', JSON.stringify(img)); setCfg_('IMG_F', JSON.stringify(files)); setCfg_('IMG_V', String(Date.now()));
  return img[key];
}
function adminImgSize(tok, key, pct) {
  adm_(tok);
  const sz = json_(cfg_('IMG_SZ'), {}); pct = Math.max(30, Math.min(200, Math.round(Number(pct) || 100)));
  if (pct === 100) delete sz[key]; else sz[key] = pct;
  setCfg_('IMG_SZ', JSON.stringify(sz)); setCfg_('IMG_V', String(Date.now()));
  return pct;
}
function adminImgReset(tok, key) {
  adm_(tok);
  const img = json_(cfg_('IMG'), {}), files = json_(cfg_('IMG_F'), {});
  if (files[key]) { try { DriveApp.getFileById(files[key]).setTrashed(true); } catch (e) {} }
  delete img[key]; delete files[key];
  setCfg_('IMG', JSON.stringify(img)); setCfg_('IMG_F', JSON.stringify(files)); setCfg_('IMG_V', String(Date.now()));
  return true;
}
function setEx_(ex) { setCfg_('EX', JSON.stringify(ex)); setCfg_('EX_V', String(Date.now())); }

function adminSettings(tok) {
  adm_(tok);
  return { quota: Number(cfg_('QUOTA_MONTH') || DEF.QUOTA_MONTH), used: quotaUsed_(), textModel: cfg_('TEXT_MODEL') || DEF.TEXT_MODEL,
    imgModel: cfg_('IMG_MODEL') || DEF.IMG_MODEL, site: cfg_('SITE_URL') || '', hasKey: !!PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') };
}
function adminSaveSettings(tok, s) {
  adm_(tok);
  if (s.quota !== undefined) { const q = Math.max(0, Math.floor(Number(s.quota) || 0)); setCfg_('QUOTA_MONTH', String(q)); }
  if (s.textModel) setCfg_('TEXT_MODEL', String(s.textModel).trim());
  if (s.imgModel) setCfg_('IMG_MODEL', String(s.imgModel).trim());
  if (s.site !== undefined) setCfg_('SITE_URL', String(s.site).trim());
  return adminSettings(tok);
}
function adminTestKey(tok) {
  adm_(tok);
  const res = gemini_(cfg_('TEXT_MODEL') || DEF.TEXT_MODEL, { contents: [{ parts: [{ text: '用一句繁體中文跟健身的人打招呼，10 個字以內。' }] }] });
  return (((res.candidates || [])[0] || {}).content || {}).parts.map(x => x.text || '').join('').trim();
}
function adminResetPin(tok, name) {
  adm_(tok);
  return locked_(() => {
    const sh = sheet_(SH.users), row = findRow_(sh, name);
    if (!row) throw new Error('找不到 ' + name);
    const salt = Utilities.getUuid();
    sh.getRange(row, 2, 1, 2).setValues([[hash_('0000', salt), '']]);
    sh.getRange(row, 6).setValue(salt);
    return name + ' 的密碼已重設為 0000';
  });
}
function adminDelete(tok, name) {
  adm_(tok);
  return locked_(() => {
    [SH.users, SH.data].forEach(n => { const sh = sheet_(n), row = findRow_(sh, name); if (row) sh.deleteRow(row); });
    return name + ' 已刪除';
  });
}

// ================= 共用 =================
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function locked_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
function clean_(s) { return String(s || '').trim().slice(0, 20); }
function json_(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }
function day_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v); }
function weekStart_(d) { const x = new Date(d); const g = (x.getDay() + 6) % 7; x.setDate(x.getDate() - g); x.setHours(0, 0, 0, 0); return x; }
function hash_(pin, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin + ':' + salt).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function sheet_(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(n);
  if (!sh) { sh = ss.insertSheet(n); sh.appendRow(HEAD[n]); sh.setFrozenRows(1); }
  return sh;
}
function rows_(n) { return sheet_(n).getDataRange().getValues().slice(1); }
function findRow_(sh, key) {
  const v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === key) return i + 1;
  return 0;
}
function cfg_(k) { const v = rows_(SH.config).find(x => x[0] === k); return v ? String(v[1]) : ''; }
function setCfg_(k, val) {
  locked_(() => { const sh = sheet_(SH.config), row = findRow_(sh, k); if (row) sh.getRange(row, 2).setValue(val); else sh.appendRow([k, val]); });
}
function folder_(sub) {
  const props = PropertiesService.getScriptProperties();
  let rootId = props.getProperty('DRIVE_ROOT'), root;
  try { root = rootId && DriveApp.getFolderById(rootId); } catch (e) { root = null; }
  if (!root) { root = DriveApp.createFolder('翹臀冒險素材'); props.setProperty('DRIVE_ROOT', root.getId()); }
  const it = root.getFoldersByName(sub);
  return it.hasNext() ? it.next() : root.createFolder(sub);
}

/** 第一次部署前可以在編輯器執行一次：建立所有工作表並要求授權（試算表、雲端硬碟、外部連線） */
function setup() {
  Object.keys(SH).forEach(k => sheet_(SH[k]));
  folder_('pets'); folder_('exercises'); folder_('assets');
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('完成！工作表與「翹臀冒險素材」資料夾都建好了。');
}
