/**
 * Alerts to Telegram, LINE (Messaging API push) and email. Tokens live in the
 * document's properties, not in a cell, so they never show up in the sheet
 * and are not copied when someone makes a copy of the file.
 */

var SETTING_KEYS = ['shopName', 'reportHour', 'email', 'telegramToken', 'telegramChatId', 'lineToken', 'lineTo'];

function getSettings() {
  var props = PropertiesService.getDocumentProperties().getProperties();
  var s = {};
  SETTING_KEYS.forEach(function (k) { s[k] = props[k] || ''; });
  if (!s.shopName) s.shopName = DEMO_SHOP;
  if (!s.reportHour) s.reportHour = '20';
  return s;
}

/** Called from Settings.html. Empty secret fields keep the saved value. */
function saveSettings(form) {
  var props = PropertiesService.getDocumentProperties();
  var hour = Number(form.reportHour);
  if (!(hour >= 0 && hour <= 23)) throw new Error('เวลาส่งสรุปต้องเป็นชั่วโมง 0-23');
  if (form.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(form.email)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
  SETTING_KEYS.forEach(function (k) {
    var v = String(form[k] || '').trim();
    var secret = k === 'telegramToken' || k === 'lineToken';
    if (secret && !v) return;
    props.setProperty(k, v);
  });
  if (findDailyTrigger()) installDailyTrigger(hour);   // keep the trigger in sync with the hour
  return settingsForForm();
}

/** Settings for the form: secrets are masked, only whether they are set is sent. */
function settingsForForm() {
  var s = getSettings();
  return {
    shopName: s.shopName, reportHour: s.reportHour, email: s.email, telegramChatId: s.telegramChatId, lineTo: s.lineTo,
    telegramSet: !!s.telegramToken, lineSet: !!s.lineToken, dailyOn: !!findDailyTrigger(),
  };
}

function clearSecret(key) {
  if (key !== 'telegramToken' && key !== 'lineToken') throw new Error('unknown key');
  PropertiesService.getDocumentProperties().deleteProperty(key);
  return settingsForForm();
}

function sendTelegram(token, chatId, text) {
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', payload: { chat_id: chatId, text: text }, muteHttpExceptions: true,
  });
  var body = JSON.parse(res.getContentText() || '{}');
  if (!body.ok) throw new Error(body.description || 'HTTP ' + res.getResponseCode());
}

function sendLine(token, to, text) {
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text.slice(0, 5000) }] }),
  });
  if (res.getResponseCode() !== 200) {
    var body = JSON.parse(res.getContentText() || '{}');
    throw new Error(body.message || 'HTTP ' + res.getResponseCode());
  }
}

/** Send to every configured channel. Returns [{channel, ok, error}] (empty when none is set up). */
function notifyAll(subject, text) {
  var s = getSettings();
  var jobs = [];
  if (s.telegramToken && s.telegramChatId) jobs.push(['Telegram', function () { sendTelegram(s.telegramToken, s.telegramChatId, text); }]);
  if (s.lineToken && s.lineTo) jobs.push(['LINE', function () { sendLine(s.lineToken, s.lineTo, text); }]);
  if (s.email) jobs.push(['อีเมล', function () { MailApp.sendEmail(s.email, subject, text); }]);
  return jobs.map(function (job) {
    try {
      job[1]();
      return { channel: job[0], ok: true };
    } catch (err) {
      console.error(job[0] + ' failed: ' + err.message);
      return { channel: job[0], ok: false, error: err.message };
    }
  });
}

function describeResults(results) {
  if (!results.length) return 'ยังไม่ได้ตั้งค่าช่องทางแจ้งเตือน (เมนู 📦 ระบบสต็อก › ⚙️ ตั้งค่าการแจ้งเตือน)';
  return results.map(function (r) { return (r.ok ? '✅ ส่งทาง ' : '❌ ส่งทาง ') + r.channel + (r.ok ? ' แล้ว' : ' ไม่สำเร็จ: ' + r.error); }).join('\n');
}

/** "ทดสอบส่ง" button in Settings.html. */
function sendTestMessage() {
  var s = getSettings();
  return describeResults(notifyAll('ทดสอบการแจ้งเตือน', '🔔 ทดสอบการแจ้งเตือนจากระบบสต็อก ' + s.shopName));
}
