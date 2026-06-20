/**
 * PTA運営支援ツール
 *
 * Googleフォーム、Googleスプレッドシート、Apps Scriptだけで動く、
 * PTA入会申込・同意記録・入金確認・会員名簿作成のための最小構成です。
 *
 * このツールは決済代行ではありません。お金の移動は扱わず、
 * 「申込」「同意」「入金確認」「会員確定」の記録を残します。
 *
 * ライセンス: MIT
 * 公開: PTA適正化推進委員会
 */

// ===== 設定：最初にこのブロックだけ各PTAで書き換えてください =====
const CONFIG = {
  ORG_NAME: 'サンプルPTA',
  CONTACT_EMAIL: 'pta@example.com',
  CONSENT_VERSION: '2026-06',
  ROSTER_SHEET: '申込・名簿',
  MEMBER_SHEET: '会員名簿',
  PAYMENT_GUIDE_TEXT: '会費のお支払い方法は、別途配布済みの案内をご確認ください。',
  LINE_CHANNEL_ACCESS_TOKEN: '',
};

// Googleフォームの質問タイトルです。
// 右側の文字列は、実際のフォームの質問文と一字一句そろえてください。
const FIELD_MAP = {
  guardianName: '保護者氏名',
  childName: 'お子さまの氏名',
  gradeClass: '学年・組',
  email: 'メールアドレス',
  consent: '重要事項に同意します',
};

// 状態。申込受付と会員確定を明確に分けます。
const STATUS_APPLIED = '申込受付（入金前）';
const STATUS_MEMBER = '会員（入金確認済）';
const STATUS_WITHDRAWN = '取下げ／無効';
const STATUS_NEEDS_REVIEW = '確認要（同意未確認）';

// 旧版から移行したシートも扱えるようにするための互換値。
const LEGACY_PENDING_VALUES = ['申込（未入金）', '未納'];
const LEGACY_MEMBER_VALUES = ['会員（入金済）', '入金済'];

const COL = {
  MEMBER_ID: 1,
  TIMESTAMP: 2,
  GUARDIAN_NAME: 3,
  CHILD_NAME: 4,
  GRADE_CLASS: 5,
  EMAIL: 6,
  CONSENT: 7,
  CONSENT_VERSION: 8,
  STATUS: 9,
  JOIN_DATE: 10,
  GUIDE_COUNT: 11,
  GUIDE_DATE: 12,
  NOTE: 13,
};
const ROSTER_COLS = 13;

// ===== 初期設定 =====
// Apps Scriptに貼り付けたあと、最初に1回だけ実行してください。
function setupPtaTool() {
  const roster = getOrCreateRoster_();
  getOrCreateMemberSheet_();
  applyRosterFormatting_(roster);
  ui_('初期設定が完了しました。次にフォーム送信トリガーを onFormSubmit に設定してください。');
}

// スプレッドシートを開いたときにメニューを追加します。
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PTAツール')
    .addItem('初期設定を実行', 'setupPtaTool')
    .addSeparator()
    .addItem('選択した行を会員として確定（入金確認）', 'confirmMembers')
    .addItem('選択した行を取下げ／無効にする', 'withdrawApplications')
    .addItem('会員名簿を書き出す（入金確認済のみ）', 'exportMembers')
    .addSeparator()
    .addItem('申込受付中の方へ支払い案内メールを送信', 'sendPaymentGuides')
    .addItem('状況の集計を表示', 'showSummary')
    .addSeparator()
    .addItem('LINEで一斉連絡（任意）', 'broadcastLinePrompt')
    .addItem('設定を確認', 'showConfig')
    .addToUi();
}

// ===== フォーム送信時に自動実行 =====
// インストール型トリガーで「スプレッドシートから」「フォーム送信時」に設定します。
function onFormSubmit(e) {
  const namedValues = e && e.namedValues ? e.namedValues : {};
  const value = function (title) {
    const v = namedValues[title];
    return v && v[0] !== undefined ? String(v[0]).trim() : '';
  };

  const roster = getOrCreateRoster_();
  const memberId = nextMemberId_();
  const consentValue = value(FIELD_MAP.consent);
  const consented = consentValue !== '';
  const status = consented ? STATUS_APPLIED : STATUS_NEEDS_REVIEW;

  const row = new Array(ROSTER_COLS).fill('');
  row[COL.MEMBER_ID - 1] = memberId;
  row[COL.TIMESTAMP - 1] = new Date();
  row[COL.GUARDIAN_NAME - 1] = value(FIELD_MAP.guardianName);
  row[COL.CHILD_NAME - 1] = value(FIELD_MAP.childName);
  row[COL.GRADE_CLASS - 1] = value(FIELD_MAP.gradeClass);
  row[COL.EMAIL - 1] = value(FIELD_MAP.email);
  row[COL.CONSENT - 1] = consented ? '同意' : '未確認';
  row[COL.CONSENT_VERSION - 1] = CONFIG.CONSENT_VERSION;
  row[COL.STATUS - 1] = status;
  row[COL.JOIN_DATE - 1] = '';
  row[COL.GUIDE_COUNT - 1] = 0;
  row[COL.GUIDE_DATE - 1] = '';
  row[COL.NOTE - 1] = consented ? '' : 'フォーム項目名または必須チェックの設定を確認';

  roster.appendRow(row);
  applyRosterFormatting_(roster);

  const email = row[COL.EMAIL - 1];
  if (email) {
    sendConfirmation_(email, memberId, row[COL.GUARDIAN_NAME - 1], consented);
  }
}

// ===== 入金確認：ここで初めて会員として確定します =====
function confirmMembers() {
  const sh = getRosterActiveSheet_();
  if (!sh) return;

  const sel = sh.getActiveRange();
  if (!sel) {
    ui_('会員として確定する行を選択してください。');
    return;
  }

  let confirmed = 0;
  let skipped = 0;
  forEachSelectedDataRow_(sel, function (rowNumber) {
    const status = String(sh.getRange(rowNumber, COL.STATUS).getValue() || '').trim();
    const consent = String(sh.getRange(rowNumber, COL.CONSENT).getValue() || '').trim();
    if (!isPending_(status) || consent !== '同意') {
      skipped++;
      return;
    }

    sh.getRange(rowNumber, COL.STATUS).setValue(STATUS_MEMBER);
    if (!sh.getRange(rowNumber, COL.JOIN_DATE).getValue()) {
      sh.getRange(rowNumber, COL.JOIN_DATE).setValue(new Date());
    }
    confirmed++;
  });

  ui_('会員として確定: ' + confirmed + '件\n対象外としてスキップ: ' + skipped + '件');
}

// 申込を取り下げた、重複、誤送信などの行を会員名簿から除外します。
function withdrawApplications() {
  const sh = getRosterActiveSheet_();
  if (!sh) return;

  const sel = sh.getActiveRange();
  if (!sel) {
    ui_('取下げ／無効にする行を選択してください。');
    return;
  }

  let updated = 0;
  forEachSelectedDataRow_(sel, function (rowNumber) {
    sh.getRange(rowNumber, COL.STATUS).setValue(STATUS_WITHDRAWN);
    updated++;
  });

  ui_('取下げ／無効に変更: ' + updated + '件');
}

// ===== 会員名簿の書き出し =====
// 入金確認済みの人だけを書き出します。
function exportMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  if (!src || src.getLastRow() < 2) {
    ui_('申込・名簿にデータがありません。');
    return;
  }

  const data = src.getRange(2, 1, src.getLastRow() - 1, ROSTER_COLS).getValues();
  const out = [[
    '会員ID',
    '保護者氏名',
    'お子さまの氏名',
    '学年・組',
    'メールアドレス',
    '会員成立日',
  ]];

  data.forEach(function (r) {
    if (isMember_(r[COL.STATUS - 1])) {
      out.push([
        r[COL.MEMBER_ID - 1],
        r[COL.GUARDIAN_NAME - 1],
        r[COL.CHILD_NAME - 1],
        r[COL.GRADE_CLASS - 1],
        r[COL.EMAIL - 1],
        r[COL.JOIN_DATE - 1],
      ]);
    }
  });

  const sh = getOrCreateMemberSheet_();
  sh.clearContents();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, out[0].length);

  ui_('会員名簿を更新しました。会員数: ' + (out.length - 1) + '件');
}

// ===== 支払い案内メール =====
// 「未納者への督促」ではなく、申込受付中の人への案内として扱います。
function sendPaymentGuides() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sh || sh.getLastRow() < 2) {
    ui_('申込・名簿にデータがありません。');
    return;
  }

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, ROSTER_COLS).getValues();
  let sent = 0;
  let skipped = 0;

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const email = String(r[COL.EMAIL - 1] || '').trim();
    if (isPending_(r[COL.STATUS - 1]) && email) {
      sendPaymentGuideMail_(email, r[COL.GUARDIAN_NAME - 1], r[COL.MEMBER_ID - 1]);
      const rowNumber = i + 2;
      sh.getRange(rowNumber, COL.GUIDE_COUNT).setValue(Number(r[COL.GUIDE_COUNT - 1] || 0) + 1);
      sh.getRange(rowNumber, COL.GUIDE_DATE).setValue(new Date());
      sent++;
      Utilities.sleep(200);
    } else {
      skipped++;
    }
  }

  ui_('支払い案内メール送信: ' + sent + '件\n対象外: ' + skipped + '件');
}

// 旧版のトリガー名を残している場合の互換用。
function sendReminders() {
  sendPaymentGuides();
}

// ===== 集計 =====
function showSummary() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sh || sh.getLastRow() < 2) {
    ui_('申込・名簿にデータがありません。');
    return;
  }

  const statuses = sh.getRange(2, COL.STATUS, sh.getLastRow() - 1, 1).getValues();
  let pending = 0;
  let member = 0;
  let withdrawn = 0;
  let review = 0;
  let other = 0;

  statuses.forEach(function (row) {
    const status = String(row[0] || '').trim();
    if (isPending_(status)) pending++;
    else if (isMember_(status)) member++;
    else if (status === STATUS_WITHDRAWN) withdrawn++;
    else if (status === STATUS_NEEDS_REVIEW) review++;
    else other++;
  });

  ui_(
    '申込受付（入金前）: ' + pending +
    '\n会員（入金確認済）: ' + member +
    '\n取下げ／無効: ' + withdrawn +
    '\n確認要: ' + review +
    '\nその他: ' + other
  );
}

function showConfig() {
  ui_(
    '団体名: ' + CONFIG.ORG_NAME +
    '\n問い合わせ先: ' + CONFIG.CONTACT_EMAIL +
    '\n同意文の版: ' + CONFIG.CONSENT_VERSION +
    '\n申込管理シート: ' + CONFIG.ROSTER_SHEET +
    '\n会員名簿シート: ' + CONFIG.MEMBER_SHEET +
    '\nLINE連携: ' + (CONFIG.LINE_CHANNEL_ACCESS_TOKEN ? '設定済' : '未設定')
  );
}

// ===== LINE一斉連絡（任意） =====
function broadcastLinePrompt() {
  const ui = SpreadsheetApp.getUi();
  if (!CONFIG.LINE_CHANNEL_ACCESS_TOKEN) {
    ui.alert('LINE連携は任意機能です。利用するには CONFIG.LINE_CHANNEL_ACCESS_TOKEN を設定してください。');
    return;
  }

  const res = ui.prompt('LINEで一斉連絡', '友だち登録者へ送るメッセージを入力してください。', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const msg = res.getResponseText().trim();
  if (!msg) {
    ui.alert('メッセージが空です。');
    return;
  }

  const code = lineBroadcast_(msg);
  ui.alert(code === 200 ? '送信しました。' : '送信に失敗しました。HTTPコード: ' + code);
}

function lineBroadcast_(message) {
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ messages: [{ type: 'text', text: message }] }),
    muteHttpExceptions: true,
  });
  return res.getResponseCode();
}

// ===== メール本文 =====
function sendConfirmation_(to, id, name, consented) {
  const body = consented
    ? (name || '保護者') + ' 様\n\n' +
      CONFIG.ORG_NAME + 'への入会申込を受け付けました。\n' +
      '受付番号: ' + id + '\n\n' +
      'この時点では「申込受付（入金前）」です。\n' +
      '会費のお支払いを本会が確認した時点で、会員登録が完了します。\n\n' +
      CONFIG.PAYMENT_GUIDE_TEXT + '\n\n' +
      '入会を取りやめる場合や、誤って送信した場合は、下記までご連絡ください。\n\n' +
      CONFIG.ORG_NAME + '\n' + CONFIG.CONTACT_EMAIL
    : (name || '保護者') + ' 様\n\n' +
      CONFIG.ORG_NAME + 'へのフォーム送信を受け付けましたが、重要事項への同意が確認できませんでした。\n' +
      '受付番号: ' + id + '\n\n' +
      '本会側で内容を確認します。必要に応じて、あらためてご連絡いたします。\n\n' +
      CONFIG.ORG_NAME + '\n' + CONFIG.CONTACT_EMAIL;

  MailApp.sendEmail(to, '【' + CONFIG.ORG_NAME + '】入会申込受付のお知らせ', body);
}

function sendPaymentGuideMail_(to, name, id) {
  const body =
    (name || '保護者') + ' 様\n\n' +
    CONFIG.ORG_NAME + 'です。\n' +
    '入会申込を受け付けていますが、会費のお支払い確認がまだ完了していません。\n' +
    '受付番号: ' + id + '\n\n' +
    CONFIG.PAYMENT_GUIDE_TEXT + '\n\n' +
    'お支払いを確認した時点で、会員登録が完了します。\n' +
    '入会を取りやめる場合、誤って申し込んだ場合、またはご不明点がある場合は、' + CONFIG.CONTACT_EMAIL + ' までご連絡ください。\n\n' +
    '※申込受付中の方へ送信しています。行き違いの場合はご容赦ください。';

  MailApp.sendEmail(to, '【' + CONFIG.ORG_NAME + '】会費のお支払い方法のご案内', body);
}

// ===== 補助関数 =====
function getOrCreateRoster_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.ROSTER_SHEET);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      '会員ID',
      '受付日時',
      '保護者氏名',
      'お子さまの氏名',
      '学年・組',
      'メールアドレス',
      '同意',
      '同意文の版',
      '状態',
      '入金確認日／会員成立日',
      '支払い案内回数',
      '支払い案内日',
      '備考',
    ]);
  }
  return sh;
}

function getOrCreateMemberSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.MEMBER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.MEMBER_SHEET);
    sh.appendRow(['会員ID', '保護者氏名', 'お子さまの氏名', '学年・組', 'メールアドレス', '会員成立日']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function applyRosterFormatting_(sh) {
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, ROSTER_COLS);

  const maxRows = Math.max(sh.getMaxRows() - 1, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS_APPLIED, STATUS_MEMBER, STATUS_WITHDRAWN, STATUS_NEEDS_REVIEW], true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, COL.STATUS, maxRows, 1).setDataValidation(statusRule);
}

function getRosterActiveSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sh.getName() !== CONFIG.ROSTER_SHEET) {
    ui_('「' + CONFIG.ROSTER_SHEET + '」シートで対象行を選択してから実行してください。');
    return null;
  }
  return sh;
}

function forEachSelectedDataRow_(range, callback) {
  const start = range.getRow();
  const end = start + range.getNumRows() - 1;
  for (let r = start; r <= end; r++) {
    if (r < 2) continue;
    callback(r);
  }
}

function nextMemberId_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getDocumentProperties();
    const n = Number(props.getProperty('MEMBER_SEQ') || '0') + 1;
    props.setProperty('MEMBER_SEQ', String(n));
    return 'A' + ('0000' + n).slice(-4);
  } finally {
    lock.releaseLock();
  }
}

function isPending_(status) {
  const v = String(status || '').trim();
  return v === STATUS_APPLIED || LEGACY_PENDING_VALUES.indexOf(v) !== -1;
}

function isMember_(status) {
  const v = String(status || '').trim();
  return v === STATUS_MEMBER || LEGACY_MEMBER_VALUES.indexOf(v) !== -1;
}

function ui_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}
