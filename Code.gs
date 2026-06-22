/**
 * ============================================================
 *  GVA EVENT CHECK-IN SYSTEM
 *  PHOEBE BEAUTY UP × Barbie 新作発表会
 * ============================================================
 *
 * 【セットアップ手順】
 *  1. スプレッドシートを開く
 *  2. 拡張機能 → Apps Script → このコードを貼り付け → 保存
 *  3. addHeaders を実行してH〜N列にヘッダーを追加
 *  4. 「デプロイ」→「新しいデプロイ」
 *     - 種類: ウェブアプリ
 *     - 実行ユーザー: 自分
 *     - アクセス: 全員
 *  5. トリガー設定:「トリガーを追加」
 *     - 関数: onFormSubmit
 *     - イベント: スプレッドシートから → フォーム送信時
 *  6. フォームにテスト送信して動作確認
 *
 * 【当日の使い方】
 *  参加者: iPhoneの純正カメラアプリでQRコードをスキャン → Safariで入場処理
 *  スタッフ: デプロイしたURLを開く（統計確認画面）
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
// 設定
// ─────────────────────────────────────────────────────────────
const SS_ID    = '1enoNq5Mv8Qiz5KRNigM6M6pMGX16kvArqsBaCDWNWm4';
const SH_NAME  = 'フォームの回答 1';  // ← シート名が違えば修正

// ★ デプロイ後にWebアプリのURLをここに貼り付けてください ★
// 例: const WEBAPP_URL = 'https://script.google.com/macros/s/XXXXXXXX/exec';
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxSqFWg2xqlfD3WFdXvJRR6ZD6fgmITobVnvEM3lC2xODHI__xM1WkMSlPcZZcLS2-h/exec';

const CC_EMAIL   = 'info@goodvibesagency.tokyo';
const FROM_EMAIL = 'info@goodvibesagency.tokyo';
const FROM_NAME  = 'GOOD VIBES AGENCY';
const EVENT_TITLE = '【PHOEBE BEAUTY UP × Barbie】新作発表会';
const EVENT_DATE  = '2026年7月30日（木）';
const EVENT_VENUE = 'or TOKYO 3F（SHIBUYA MIYASHITA PARK）';

// 列番号（1始まり）
const C = {
  TIMESTAMP:    1,   // A
  NAME:         2,   // B
  EMAIL:        3,   // C
  SNS_URL:      4,   // D
  SNS_TYPE:     5,   // E
  TIMESLOT:     6,   // F
  COMPANION:    7,   // G
  COMP_NAME:    8,   // H ← お連れさまの名前（フォーム追加質問）
  UID:          9,   // I ← GAS追加
  CID:          10,  // J ← GAS追加
  MAIL_SENT:    11,  // K ← GAS追加
  CHECKED_IN:   12,  // L ← GAS追加
  CHECKIN_TIME: 13,  // M ← GAS追加
  COMP_IN:      14,  // N ← GAS追加
  COMP_TIME:    15,  // O ← GAS追加
};

// ─────────────────────────────────────────────────────────────
// ユーティリティ
// ─────────────────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(SH_NAME);
}

function genId() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const rand6 = Array.from({length: 6}, () => chars[Math.floor(Math.random() * chars.length)]).join('');
  return rand6 + '-' + Date.now().toString(36).toUpperCase().slice(-4);
}

function qrImageUrl(data, size) {
  size = size || 280;
  return 'https://api.qrserver.com/v1/create-qr-code/?size=' + size + 'x' + size
       + '&data=' + encodeURIComponent(data);
}

function getWebAppUrl() {
  if (WEBAPP_URL) return WEBAPP_URL;
  // WEBAPP_URL未設定の場合は自動取得を試みる（デプロイ済み環境のみ動作）
  try { return ScriptApp.getService().getUrl(); } catch(e) { return ''; }
}

function nowJST() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss');
}

/** H〜N列にヘッダーを追加（初回のみ手動実行） */
function addHeaders() {
  const sh = getSheet();
  // J列(10)からGAS追加列のヘッダーを設定
  ['本人ID','同伴者ID','メール送信','本人入場','本人入場時刻','同伴者入場','同伴者入場時刻']
    .forEach((h, i) => sh.getRange(1, C.UID + i).setValue(h));
  SpreadsheetApp.flush();
  Logger.log('ヘッダー追加完了（J列から）');
}

// ─────────────────────────────────────────────────────────────
// フォーム送信トリガー
// ─────────────────────────────────────────────────────────────
function onFormSubmit(e) {
  const sh  = getSheet();
  const row = e.range.getRow();
  const v   = sh.getRange(row, 1, 1, 7).getValues()[0];

  const name     = v[C.NAME - 1];
  const email    = v[C.EMAIL - 1];
  const timeslot = v[C.TIMESLOT - 1];
  const hasComp  = String(v[C.COMPANION - 1]).includes('お連れ様');

  const uid = genId();
  const cid = hasComp ? genId() : '';

  sh.getRange(row, C.UID).setValue(uid);
  sh.getRange(row, C.CID).setValue(cid);
  sh.getRange(row, C.MAIL_SENT).setValue('送信中...');
  sh.getRange(row, C.CHECKED_IN).setValue('未');
  sh.getRange(row, C.CHECKIN_TIME).setValue('');
  sh.getRange(row, C.COMP_IN).setValue(hasComp ? '未' : '—');
  sh.getRange(row, C.COMP_TIME).setValue('');
  SpreadsheetApp.flush();

  // WebアプリURLが未設定の場合は警告して終了
  if (!getWebAppUrl()) {
    sh.getRange(row, C.MAIL_SENT).setValue('❌ WEBAPP_URLが未設定です');
    console.error('WEBAPP_URLを設定してください');
    return;
  }

  try {
    sendInvitationMail(email, name, timeslot, uid, cid, hasComp);
    sh.getRange(row, C.MAIL_SENT).setValue('✅ 送信済み');
  } catch(err) {
    sh.getRange(row, C.MAIL_SENT).setValue('❌ ' + err.message);
    console.error(err);
  }
}

// ─────────────────────────────────────────────────────────────
// 招待メール送信
// ─────────────────────────────────────────────────────────────
function sendInvitationMail(to, name, timeslot, uid, cid, hasComp) {
  const base = getWebAppUrl();

  // QRコードに埋め込むURL = WebアプリURL?id=xxx&type=main
  // → iPhoneの純正カメラでスキャンするとSafariで自動的に開く
  const myQr   = qrImageUrl(base + '?id=' + uid + '&type=main', 280);
  const compQr = hasComp ? qrImageUrl(base + '?id=' + cid + '&type=comp', 240) : null;

  const compSection = hasComp ? `
    <div style="margin-top:24px;padding:20px;background:#fff0f5;border-radius:12px;
                text-align:center;border:1px dashed #ffb6c1;">
      <p style="font-size:13px;color:#cc5577;font-weight:700;margin:0 0 12px;">
        お連れ様 専用QRコード
      </p>
      <img src="${compQr}" width="180" height="180"
           style="border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.1);"/>
      <p style="font-size:11px;color:#bbb;margin:10px 0 0;">
        ※こちらはお連れ様の入場専用です
      </p>
    </div>` : '';

  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#fff5f8;font-family:-apple-system,Helvetica Neue,sans-serif;">
<div style="max-width:480px;margin:0 auto;padding:20px;">

  <div style="background:linear-gradient(135deg,#ff69b4,#e91e8c);border-radius:16px;
              padding:28px;text-align:center;margin-bottom:20px;">
    <p style="color:rgba(255,255,255,.7);font-size:10px;letter-spacing:3px;margin:0 0 8px;">
      GOOD VIBES AGENCY
    </p>
    <h1 style="color:#fff;font-size:18px;margin:0;line-height:1.5;font-weight:800;">
      PHOEBE BEAUTY UP × Barbie<br>新作発表会 ご招待状
    </h1>
    <p style="color:rgba(255,255,255,.8);font-size:12px;margin:10px 0 0;">
      ${EVENT_DATE} &nbsp;|&nbsp; ${EVENT_VENUE}
    </p>
  </div>

  <div style="background:#fff;border-radius:16px;padding:24px;margin-bottom:20px;">
    <p style="font-size:15px;color:#333;margin:0 0 14px;">
      <strong>${name}</strong> 様
    </p>
    <p style="font-size:13px;color:#555;line-height:1.8;margin:0 0 20px;">
      この度はご登録いただき、誠にありがとうございます。<br>
      当日は下記のQRコードを受付スタッフにご提示ください。
    </p>

    <div style="background:#fafafa;border-radius:12px;padding:16px;margin-bottom:20px;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr>
          <td style="padding:7px 0;color:#999;white-space:nowrap;">日時</td>
          <td style="padding:7px 0;padding-left:12px;color:#333;font-weight:600;">${EVENT_DATE}</td>
        </tr>
        <tr>
          <td style="padding:7px 0;color:#999;white-space:nowrap;">会場</td>
          <td style="padding:7px 0;padding-left:12px;color:#333;">${EVENT_VENUE}</td>
        </tr>
        <tr>
          <td style="padding:7px 0;color:#999;white-space:nowrap;">入場時間帯</td>
          <td style="padding:7px 0;padding-left:12px;color:#e91e8c;font-weight:700;font-size:14px;">${timeslot}</td>
        </tr>
      </table>
    </div>

    <div style="text-align:center;padding:22px;background:#fff5f8;border-radius:12px;
                border:1px solid #ffd6e7;">
      <p style="font-size:13px;color:#333;font-weight:700;margin:0 0 14px;">
        あなたの入場QRコード
      </p>
      <img src="${myQr}" width="200" height="200"
           style="border-radius:10px;box-shadow:0 2px 12px rgba(233,30,140,.15);"/>
      <p style="font-size:11px;color:#aaa;margin:12px 0 0;">
        受付でiPhoneカメラをQRにかざしてください
      </p>
    </div>

    ${compSection}

    <div style="margin-top:20px;padding:14px;background:#fffbeb;border-radius:10px;
                border-left:3px solid #fbbf24;">
      <p style="font-size:12px;color:#92400e;margin:0;line-height:1.7;">
        ※ <strong>トークショー（18:30〜）</strong>をご覧になりたい方は、<br>
        ご登録の入場時間帯（18:15以前）にお越しください。
      </p>
    </div>
  </div>

  <div style="text-align:center;padding:10px;">
    <p style="font-size:11px;color:#ccc;margin:0;">GOOD VIBES AGENCY</p>
    <p style="font-size:11px;color:#ccc;margin:4px 0 0;">info@goodvibesagency.tokyo</p>
  </div>

</div>
</body>
</html>`;

  GmailApp.sendEmail(to, '【ご招待】' + EVENT_TITLE, 'HTMLメール対応のアプリでご確認ください。', {
    htmlBody: html,
    cc:       CC_EMAIL,
    name:     FROM_NAME,
    replyTo:  FROM_EMAIL,
  });
}

// ─────────────────────────────────────────────────────────────
// Webアプリ エントリポイント
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  const id   = e.parameter.id;
  const type = e.parameter.type || 'main';
  const api  = e.parameter.api;

  // デバッグ用: スプレッドシートの実際のIDを返す（確認後削除）
  if (api === 'debug') {
    const sh = getSheet();
    const data = sh.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        row: i + 1,
        name: data[i][C.NAME - 1],
        uid: '[' + String(data[i][C.UID - 1]) + ']',   // 前後の空白も見える
        cid: '[' + String(data[i][C.CID - 1]) + ']',
        uid_len: String(data[i][C.UID - 1]).length,
      });
    }
    return ContentService.createTextOutput(JSON.stringify(result, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // GitHub PagesからのAPI呼び出し → JSON返却
  if (api === '1') {
    var result = id ? processCheckin(id, type) : getStats();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // QRスキャン後の結果ページ（iPhoneカメラ経由）
  if (id) {
    const result = processCheckin(id, type);
    return HtmlService.createHtmlOutput(getCheckinResultHtml(result))
      .setTitle('GVA 受付')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // パラメータなし → シンプルな統計ページ
  return HtmlService.createHtmlOutput(getStaffHtml())
    .setTitle('GVA 受付管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────────────────────
// チェックイン処理
// ─────────────────────────────────────────────────────────────
function processCheckin(id, type) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const uid    = row[C.UID - 1];
    const cid    = row[C.CID - 1];
    const name   = row[C.NAME - 1];
    const slot   = row[C.TIMESLOT - 1];
    const snsUrl = row[C.SNS_URL - 1] || '';
    const rowNum = i + 1;

    if (type === 'main' && String(uid).trim() === String(id).trim()) {
      if (row[C.CHECKED_IN - 1] === '✅ 入場済み') {
        return { ok: 'already', name: name, slot: slot, time: row[C.CHECKIN_TIME - 1], snsUrl: snsUrl };
      }
      const t = nowJST();
      sh.getRange(rowNum, C.CHECKED_IN).setValue('✅ 入場済み');
      sh.getRange(rowNum, C.CHECKIN_TIME).setValue(t);
      SpreadsheetApp.flush();
      return { ok: 'success', name: name, slot: slot, time: t, snsUrl: snsUrl };
    }

    if (type === 'comp' && String(cid).trim() === String(id).trim()) {
      const label = name + ' さんのお連れ様';
      if (row[C.COMP_IN - 1] === '✅ 入場済み') {
        return { ok: 'already', name: label, slot: slot, time: row[C.COMP_TIME - 1], snsUrl: snsUrl };
      }
      const t = nowJST();
      sh.getRange(rowNum, C.COMP_IN).setValue('✅ 入場済み');
      sh.getRange(rowNum, C.COMP_TIME).setValue(t);
      SpreadsheetApp.flush();
      return { ok: 'success', name: label, slot: slot, time: t, snsUrl: snsUrl };
    }
  }
  return { ok: 'notfound' };
}

// ─────────────────────────────────────────────────────────────
// チェックイン結果ページ（QRスキャン後にiPhoneに表示される画面）
// ─────────────────────────────────────────────────────────────
function getCheckinResultHtml(r) {
  var ico, title, borderColor, bgColor;
  if (r.ok === 'success') {
    ico = '✅'; title = '入場OK'; borderColor = '#00cc66'; bgColor = '#0a2e1a';
  } else if (r.ok === 'already') {
    ico = '⚠️'; title = '入場済み'; borderColor = '#ffcc00'; bgColor = '#2e2200';
  } else {
    ico = '❌'; title = 'QRコード不明'; borderColor = '#ff4444'; bgColor = '#2e0a0a';
  }

  const nameHtml   = r.name   ? '<div class="nm">'  + r.name  + '</div>' : '';
  const slotHtml   = r.slot   ? '<div class="sl">'  + r.slot  + '</div>' : '';
  const timeHtml   = r.time   ? '<div class="tm">入場時刻: ' + r.time + '</div>' : '';
  const snsHtml    = r.snsUrl ? '<div class="sns">' + r.snsUrl + '</div>' : '';
  const subHtml    = r.ok === 'notfound' ? '<div class="sub">スタッフに確認してください</div>' : '';

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
  <title>GVA 受付</title>
  <style>
    * { box-sizing:border-box; margin:0; padding:0; }
    body {
      background:#0f0f1a; color:#fff;
      font-family:-apple-system,Helvetica,sans-serif;
      min-height:100vh; display:flex; flex-direction:column;
      align-items:center; justify-content:center; padding:24px;
    }
    .card {
      background:${bgColor}; border:2px solid ${borderColor};
      border-radius:24px; padding:36px 28px;
      max-width:360px; width:100%; text-align:center;
    }
    .ico  { font-size:64px; margin-bottom:12px; }
    .ttl  { font-size:28px; font-weight:800; color:${borderColor}; margin-bottom:16px; }
    .nm   { font-size:24px; font-weight:800; color:#fff; margin-bottom:8px; word-break:break-all; }
    .sl   { font-size:14px; color:#aaa; margin-bottom:6px; }
    .tm   { font-size:13px; color:#666; margin-bottom:4px; }
    .sns  { font-size:12px; color:#ff69b4; margin-top:14px; word-break:break-all;
            padding:10px 14px; background:rgba(255,105,180,.1); border-radius:10px; }
    .sub  { font-size:14px; color:#888; margin-top:8px; }
    .stamp {
      margin-top:20px; padding:10px 24px;
      background:${borderColor}; color:#000;
      font-size:13px; font-weight:700; border-radius:30px;
      display:inline-block;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="ico">${ico}</div>
    <div class="ttl">${title}</div>
    ${nameHtml}
    ${slotHtml}
    ${timeHtml}
    ${snsHtml}
    ${subHtml}
    ${r.ok !== 'notfound' ? '<div class="stamp">確認済み</div>' : ''}
  </div>
</body>
</html>`;
}

// ─────────────────────────────────────────────────────────────
// スタッフ用管理・スキャン画面
// ─────────────────────────────────────────────────────────────
function getStaffHtml() {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  let total = 0, cin = 0, totalComp = 0, cinComp = 0;
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[C.UID - 1]) continue;
    total++;
    const checked = row[C.CHECKED_IN - 1] === '✅ 入場済み';
    if (checked) cin++;
    const cs = row[C.COMP_IN - 1];
    if (cs === '未' || cs === '✅ 入場済み') {
      totalComp++;
      if (cs === '✅ 入場済み') cinComp++;
    }
    rows.push({
      name:    row[C.NAME - 1],
      slot:    String(row[C.TIMESLOT - 1]).replace('（トークショー前）','').replace('（※トークショー開催後）','').replace('(※最終入場 20:15)','').trim(),
      checked: checked,
      time:    row[C.CHECKIN_TIME - 1] || '',
    });
  }

  const rowsHtml = rows.map(function(r) {
    return '<tr>'
      + '<td>' + r.name + '</td>'
      + '<td>' + r.slot + '</td>'
      + '<td style="color:' + (r.checked ? '#00cc66' : '#666') + '">' + (r.checked ? 'OK ' + r.time : '未') + '</td>'
      + '</tr>';
  }).join('');

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>GVA 受付管理</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{background:#0f0f1a;color:#fff;font-family:-apple-system,Helvetica,sans-serif}
    .hd{background:linear-gradient(135deg,#ff69b4,#e91e8c);padding:16px 20px;text-align:center}
    .hd h1{font-size:17px;font-weight:800}
    .stats{display:flex;background:#16213e;padding:12px 10px;gap:8px}
    .stat{flex:1;background:#0d0d2e;border-radius:10px;padding:10px 6px;text-align:center}
    .stat .n{font-size:22px;font-weight:800;color:#ff69b4}
    .stat .l{font-size:10px;color:#555;margin-top:3px}
    .note{padding:10px 16px;background:#1a1a2e;font-size:12px;color:#555;text-align:center}
    table{width:100%;border-collapse:collapse;font-size:12px}
    th{background:#16213e;padding:8px 6px;text-align:left;color:#666}
    td{padding:8px 6px;border-bottom:1px solid #111;color:#ccc}
    .btn{display:block;text-align:center;padding:14px;background:#ff69b4;color:#fff;font-weight:700;font-size:14px;border:none;width:100%;cursor:pointer}
  </style>
</head>
<body>
  <div class="hd"><h1>GVA 受付管理</h1><p>PHOEBE BEAUTY UP x Barbie 新作発表会</p></div>
  <div class="stats">
    <div class="stat"><div class="n">${total}</div><div class="l">総登録者</div></div>
    <div class="stat"><div class="n">${cin}</div><div class="l">入場済み</div></div>
    <div class="stat"><div class="n">${total - cin}</div><div class="l">未入場</div></div>
    <div class="stat"><div class="n">${cinComp}/${totalComp}</div><div class="l">同伴</div></div>
  </div>
  <div class="note">スキャンはGitHub Pagesのページで行ってください</div>
  <table>
    <tr><th>名前</th><th>時間帯</th><th>入場</th></tr>
    ${rowsHtml}
  </table>
  <button class="btn" onclick="location.reload()">再読み込み</button>
</body>
</html>`;
}

// ─────────────────────────────────────────────────────────────
// スタッフ画面からのQRスキャン処理（URLを受け取りIDを抽出）
// ─────────────────────────────────────────────────────────────
function processCheckinByUrl(qrData) {
  // QRの内容はWebアプリURL: .../exec?id=XXXX&type=main
  try {
    var url = qrData;
    var idMatch   = url.match(/[?&]id=([^&]+)/);
    var typeMatch = url.match(/[?&]type=([^&]+)/);
    var id   = idMatch   ? decodeURIComponent(idMatch[1])   : '';
    var type = typeMatch ? decodeURIComponent(typeMatch[1]) : 'main';
    if (!id) return { ok: 'notfound' };
    return processCheckin(id, type);
  } catch(e) {
    return { ok: 'notfound' };
  }
}

// ─────────────────────────────────────────────────────────────
// 統計取得（後方互換用・今は未使用）
// ─────────────────────────────────────────────────────────────
function getStats() {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  let total = 0, cin = 0, totalComp = 0, cinComp = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][C.UID - 1]) continue;
    total++;
    if (data[i][C.CHECKED_IN - 1] === '✅ 入場済み') cin++;
    const cs = data[i][C.COMP_IN - 1];
    if (cs === '未' || cs === '✅ 入場済み') { totalComp++; if (cs === '✅ 入場済み') cinComp++; }
  }
  return { total, cin, totalComp, cinComp };
}
