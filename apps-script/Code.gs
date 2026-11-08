/**
 * 東浪工作室 預約後端（Google Apps Script）
 *
 * GET  ?action=busy&from=YYYY-MM-DD&days=N  → 回傳忙碌時段（只有開始/結束時間，不含任何內容）
 * POST {name, phone, date, time, mins, city, venue, km, note}
 *      → 在「4 工作/家教」建立【待確認】活動，並寄信通知
 *
 * 部署：以「我」的身分執行，存取權「任何人」。步驟見 README.md。
 */

const CONFIG = {
  CALENDAR_ID: 's4ejag0kgq9f9thv8ogbt8fuos@group.calendar.google.com',
  TZ: 'Asia/Taipei',
  BASE_PRICE: 2300,
  BASE_MIN: 100,
  EXT_PRICE: 600,          // 每延長 30 分
  ALLOWED_MINS: [100, 130, 160],
  EVENT_EXTRA_MIN: 20,     // 日曆活動比課程多留的時間（100 分課程 → 2 小時活動）
  BUFFER_MIN: 30,          // 前後預約之間至少間隔
  MIN_LEAD_HOURS: 3,       // 最快幾小時後可約
  MAX_DAYS_AHEAD: 60,
  MAX_PER_DAY: 30,         // 每天網站最多收幾筆（防灌單）
  NOTIFY_EMAIL: ''         // 留空 = 寄給部署這支程式的帳號
};

const CITY_CODE = { '台東': 'T', '高雄': 'K', '台南': 'N' };

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'busy') return json_({ ok: true, service: 'eastwave' });
  const days = Math.min(Math.max(parseInt(p.days, 10) || 14, 1), CONFIG.MAX_DAYS_AHEAD);
  const from = parseDate_(p.from) || startOfToday_();
  const to = new Date(from.getTime() + (days + 1) * 864e5);
  const busy = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID)
    .getEvents(from, to)
    .filter(ev => !ev.isAllDayEvent())
    .map(ev => [ev.getStartTime().toISOString(), ev.getEndTime().toISOString()]);
  return json_({ ok: true, busy: busy, bufferMin: CONFIG.BUFFER_MIN, leadHours: CONFIG.MIN_LEAD_HOURS });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  const name = String(b.name || '').trim().slice(0, 20);
  const phone = String(b.phone || '').replace(/\D/g, '');
  const note = String(b.note || '').trim().slice(0, 300);
  const city = String(b.city || '');
  const venue = String(b.venue || '').slice(0, 20);
  const mins = parseInt(b.mins, 10);
  const km = Math.max(0, Math.ceil(Number(b.km) || 0));
  const start = parseDateTime_(b.date, b.time);

  if (!name) return json_({ ok: false, error: 'name' });
  if (!/^09\d{8}$/.test(phone)) return json_({ ok: false, error: 'phone' });
  if (CONFIG.ALLOWED_MINS.indexOf(mins) < 0) return json_({ ok: false, error: 'mins' });
  if (!start) return json_({ ok: false, error: 'time' });
  const now = new Date();
  if (start.getTime() < now.getTime() + CONFIG.MIN_LEAD_HOURS * 36e5) return json_({ ok: false, error: 'too_soon' });
  if (start.getTime() > now.getTime() + CONFIG.MAX_DAYS_AHEAD * 864e5) return json_({ ok: false, error: 'too_far' });

  // 同一支手機 10 分鐘內只能送一次
  const cache = CacheService.getScriptCache();
  if (cache.get('p' + phone)) return json_({ ok: false, error: 'rate' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy_server' });
  try {
    const props = PropertiesService.getScriptProperties();
    const dayKey = 'n' + Utilities.formatDate(now, CONFIG.TZ, 'yyyyMMdd');
    const count = parseInt(props.getProperty(dayKey) || '0', 10);
    if (count >= CONFIG.MAX_PER_DAY) return json_({ ok: false, error: 'rate' });

    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    const end = new Date(start.getTime() + (mins + CONFIG.EVENT_EXTRA_MIN) * 6e4);
    const buf = CONFIG.BUFFER_MIN * 6e4;
    const clash = cal.getEvents(new Date(start.getTime() - buf), new Date(end.getTime() + buf))
      .some(ev => !ev.isAllDayEvent());
    if (clash) return json_({ ok: false, error: 'taken' });

    // 新客 / 舊客：用手機號碼找過去兩年的活動
    const past = cal.getEvents(new Date(now.getTime() - 730 * 864e5), now, { search: phone });
    const tag = past.length ? '#O' : '#N';

    const isOut = venue === '外出服務';
    const ext = (mins - CONFIG.BASE_MIN) / 30 * CONFIG.EXT_PRICE;
    const night = nightFee_(start);
    const outFee = isOut ? (km <= 5 ? 200 : 200 + (km - 5) * 50) : 0;
    const total = CONFIG.BASE_PRICE + ext + night + outFee;

    const code = CITY_CODE[city] ? 'M' + CITY_CODE[city] : 'M' + city;
    const place = isOut ? '外出' : (city === '台東' ? '' : venue.replace('行程場地', ''));
    const title = '【待確認】' + tag + ' ' + code + place + '-' + total;

    const when = Utilities.formatDate(start, CONFIG.TZ, 'M/d HH:mm');
    const desc = [
      '貴姓＋手機：' + name + ' ' + phone,
      '日期時間：' + when,
      '方案：招牌放鬆課程 ' + mins + ' 分鐘',
      '地點：' + city + ' ' + venue + (isOut ? '（約 ' + km + ' 公里）' : ''),
      '預估費用：' + total + '（延長 ' + ext + '／夜間 ' + night + '／外出 ' + outFee + '）',
      note ? '備註：' + note : null,
      '',
      '來源：網站預約 ' + Utilities.formatDate(now, CONFIG.TZ, 'yyyy/M/d HH:mm'),
      '確認後請把標題的【待確認】刪掉；不成立就刪除這筆活動。'
    ].filter(function (s) { return s !== null; }).join('\n');

    const ev = cal.createEvent(title, start, end, { description: desc });
    try { ev.setColor(CalendarApp.EventColor.GRAY); } catch (err) {}

    props.setProperty(dayKey, String(count + 1));
    cache.put('p' + phone, '1', 600);

    const to = CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(to, '東浪新預約 ' + tag + ' ' + when + ' ' + name, title + '\n\n' + desc);

    return json_({ ok: true, tag: tag, total: total });
  } finally {
    lock.releaseLock();
  }
}

function nightFee_(d) {
  const h = parseInt(Utilities.formatDate(d, CONFIG.TZ, 'H'), 10);
  if (h >= 1 && h < 6) return 400;
  if (h >= 23 || h === 0) return 200;
  return 0;
}

function parseDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) return null;
  return new Date(s + 'T00:00:00+08:00');
}

function parseDateTime_(d, t) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d || '')) || !/^\d{2}:\d{2}$/.test(String(t || ''))) return null;
  const x = new Date(d + 'T' + t + ':00+08:00');
  return isNaN(x.getTime()) ? null : x;
}

function startOfToday_() {
  return parseDate_(Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 部署前先在編輯器按「執行」跑一次，讓 Google 要求日曆與 Gmail 權限。 */
function authorize() {
  CalendarApp.getCalendarById(CONFIG.CALENDAR_ID).getName();
  MailApp.getRemainingDailyQuota();
}
