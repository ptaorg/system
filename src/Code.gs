/**
 * PTA入会申込・会員名簿管理ツール
 *
 * Googleフォーム、Googleスプレッドシート、Apps Scriptだけで動く、
 * PTAの入会申込、重要事項同意、入金確認、会員名簿作成のための最小構成です。
 *
 * このツールは決済代行ではありません。
 * 学校名簿の取り込みも前提にしていません。
 * PTA自身がフォームで取得した申込情報だけを、申込受付、入金確認、会員確定に分けて記録します。
 *
 * 公開: PTA適正化推進委員会
 * License: MIT
 */

// ============================================================================
// 1. 初期設定：まずこのブロックだけを各PTAに合わせて書き換えてください
// ============================================================================
const CONFIG = {
  // 団体名。メール件名・本文・設定表示に使います。
  ORG_NAME: 'サンプルPTA',

  // 問い合わせ先。申込者への案内文に表示します。
  CONTACT_EMAIL: 'pta@example.com',

  // 同意文・説明文の版。説明文を改訂したら年月などを更新してください。
  CONSENT_VERSION: '2026-06',

  // 申込管理用シート名。通常は変更不要です。
  ROSTER_SHEET: '申込・名簿',

  // 会員として確定した人だけを書き出すシート名。通常は変更不要です。
  MEMBER_SHEET: '会員名簿',

  // 初期値は false です。テスト完了後、メール送信を使う場合だけ true にしてください。
  EMAIL_ENABLED: false,

  // 支払い案内本文。振込先や集金方法は各PTAの実情に合わせて書き換えてください。
  PAYMENT_GUIDE_TEXT: '会費のお支払い方法は、別途配布済みの案内をご確認ください。',

  // 加入意思確認欄を作る場合、加入申込として扱う選択肢の文言をここに合わせます。
  APPLY_VALUE: 'PTAへの加入を申し込みます',
};

// Googleフォームの質問タイトルです。
// 右側の文字列は、実際のGoogleフォームの質問文と一字一句そろえてください。
// 加入意思確認欄をフォームに置かない場合でも、このままで動きます。
const FIELD_MAP = {
  guardianName: '保護者氏名',
  childName: '児童・生徒氏名',
  gradeClass: '学年・組',
  email: 'メールアドレス',
  joinIntent: '加入意思の確認',
  consent: '重要事項に同意します',
};

// ============================================================================
// 2. 状態定義：申込受付と会員確定を分けて扱います
// ============================================================================
const STATUS_APPLIED = '申込受付（入金前）';
const STATUS_MEMBER = '会員（入金確認済）';
const STATUS_NOT_APPLIED = '申込なし（記録のみ）';
const STATUS_WITHDRAWN = '取下げ／無効';
const STATUS_NEEDS_REVIEW = '確認要';

// 旧版から移行したシートも扱えるようにするための互換値。
const LEGACY_PENDING_VALUES = ['申込（未入金）', '未納'];
const LEGACY_MEMBER_VALUES = ['会員（入金済）', '入金済'];

const COL = {
  APPLICATION_ID: 1,
  TIMESTAMP: 2,
  GUARDIAN_NAME: 3,
  CHILD_NAME: 4,
  GRADE_CLASS: 5,
  EMAIL: 6,
  JOIN_INTENT: 7,
  CONSENT: 8,
  CONSENT_VERSION: 9,
  STATUS: 10,
  JOIN_DATE: 11,
  GUIDE_COUNT: 12,
  GUIDE_DATE: 13,
  NOTE: 14,
};
const ROSTER_COLS = 14;

// ============================================================================
// 3. 初期設定とメニュー
// ============================================================================

/**
 * Apps Scriptに貼り付けたあと、最初に1回だけ実行してください。
 */
function setupPtaTool() {
  const roster = getOrCreateRoster_();
  const members = getOrCreateMemberSheet_();
  const guide = getOrCreateGuideSheet_();
  applyRosterFormatting_(roster);
  applyMemberFormatting_(members);
  guide.autoResizeColumns(1, 2);

  alert_(
    '初期設定が完了しました。\n\n' +
    '次に、Apps Scriptのトリガー画面で onFormSubmit を「スプレッドシートから」「フォーム送信時」に設定してください。\n\n' +
    'テスト送信が終わるまでは CONFIG.EMAIL_ENABLED は false のままにしてください。'
  );
}

/**
 * スプレッドシートを開いたときに専用メニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PTA入会管理')
    .addItem('初期設定を実行', 'setupPtaTool')
    .addSeparator()
    .addItem('選択した行を会員として確定（入金確認）', 'confirmMembers')
    .addItem('選択した行を取下げ／無効にする', 'withdrawApplications')
    .addItem('選択した行を申込なしにする', 'markAsNotApplied')
    .addSeparator()
    .addItem('会員名簿を書き出す（入金確認済のみ）', 'exportMembers')
    .addItem('申込受付中の方へ支払い案内メールを送信', 'sendPaymentGuides')
    .addSeparator()
    .addItem('状況の集計を表示', 'showSummary')
    .addItem('設定を確認', 'showConfig')
    .addToUi();
}

// ============================================================================
// 4. フォーム送信時の処理
// ============================================================================

/**
 * インストール型トリガーで「スプレッドシートから」「フォーム送信時」に設定します。
 */
function onFormSubmit(e) {
  const namedValues = e && e.namedValues ? e.namedValues : {};
  const value = function (title) {
    const v = namedValues[title];
    return v && v[0] !== undefined ? String(v[0]).trim() : '';
  };
  const hasField = function (title) {
    return Object.prototype.hasOwnProperty.call(namedValues, title);
  };

  const roster = getOrCreateRoster_();
  const applicationId = nextApplicationId_();

  const joinIntent = value(FIELD_MAP.joinIntent);
  const consentValue = value(FIELD_MAP.consent);
  const consented = consentValue !== '';

  // 加入意思確認欄がないフォームでは、送信自体を申込として扱います。
  // 加入意思確認欄がある場合は、CONFIG.APPLY_VALUE と一致または含む回答だけを申込扱いにします。
  const hasJoinIntentQuestion = FIELD_MAP.joinIntent && hasField(FIELD_MAP.joinIntent);
  const wantsToApply = !hasJoinIntentQuestion || joinIntent === CONFIG.APPLY_VALUE || joinIntent.indexOf(CONFIG.APPLY_VALUE) !== -1;

  let status = STATUS_APPLIED;
  let note = '';
  if (!wantsToApply) {
    status = STATUS_NOT_APPLIED;
    note = '加入申込として扱わない回答';
  } else if (!consented) {
    status = STATUS_NEEDS_REVIEW;
    note = '重要事項への同意が確認できないため要確認';
  }

  const row = new Array(ROSTER_COLS).fill('');
  row[COL.APPLICATION_ID - 1] = applicationId;
  row[COL.TIMESTAMP - 1] = new Date();
  row[COL.GUARDIAN_NAME - 1] = value(FIELD_MAP.guardianName);
  row[COL.CHILD_NAME - 1] = value(FIELD_MAP.childName);
  row[COL.GRADE_CLASS - 1] = value(FIELD_MAP.gradeClass);
  row[COL.EMAIL - 1] = value(FIELD_MAP.email);
  row[COL.JOIN_INTENT - 1] = hasJoinIntentQuestion ? joinIntent : CONFIG.APPLY_VALUE;
  row[COL.CONSENT - 1] = consented ? '同意' : '未確認';
  row[COL.CONSENT_VERSION - 1] = CONFIG.CONSENT_VERSION;
  row[COL.STATUS - 1] = status;
  row[COL.JOIN_DATE - 1] = '';
  row[COL.GUIDE_COUNT - 1] = 0;
  row[COL.GUIDE_DATE - 1] = '';
  row[COL.NOTE - 1] = note;

  roster.appendRow(row);
  applyRosterFormatting_(roster);

  if (CONFIG.EMAIL_ENABLED && row[COL.EMAIL - 1]) {
    sendReceptionMail_(row[COL.EMAIL - 1], row[COL.GUARDIAN_NAME - 1], applicationId, status);
  }
}

// ============================================================================
// 5. 入金確認・状態変更
// ============================================================================

/**
 * 入金を確認した行だけ、ここで初めて会員として確定します。
 */
function confirmMembers() {
  const sh = getRosterActiveSheet_();
  if (!sh) return;

  const sel = sh.getActiveRange();
  if (!sel) {
    alert_('会員として確定する行を選択してください。');
    return;
  }

  let confirmed = 0;
  let skipped = 0;
  forEachSelectedDataRow_(sel, function (rowNumber) {
    const status = String(sh.getRange(rowNumber, COL.STATUS).getValue() || '').trim();
    const consent = String(sh.getRange(rowNumber, COL.CONSENT).getValue() || '').trim();
    const email = String(sh.getRange(rowNumber, COL.EMAIL).getValue() || '').trim();

    if (!isPending_(status) || consent !== '同意') {
      skipped++;
      return;
    }

    sh.getRange(rowNumber, COL.STATUS).setValue(STATUS_MEMBER);
    if (!sh.getRange(rowNumber, COL.JOIN_DATE).getValue()) {
      sh.getRange(rowNumber, COL.JOIN_DATE).setValue(new Date());
    }
    confirmed++;

    if (CONFIG.EMAIL_ENABLED && email) {
      sendMemberConfirmedMail_(email, sh.getRange(rowNumber, COL.GUARDIAN_NAME).getValue(), sh.getRange(rowNumber, COL.APPLICATION_ID).getValue());
    }
  });

  alert_('会員として確定: ' + confirmed + '件\n対象外としてスキップ: ' + skipped + '件');
}

/**
 * 取下げ、誤送信、重複申込などを会員名簿から除外します。
 */
function withdrawApplications() {
  updateSelectedStatuses_(STATUS_WITHDRAWN, '取下げ／無効に変更');
}

/**
 * 加入申込ではない回答として扱います。
 */
function markAsNotApplied() {
  updateSelectedStatuses_(STATUS_NOT_APPLIED, '申込なしに変更');
}

function updateSelectedStatuses_(newStatus, messageLabel) {
  const sh = getRosterActiveSheet_();
  if (!sh) return;

  const sel = sh.getActiveRange();
  if (!sel) {
    alert_('状態を変更する行を選択してください。');
    return;
  }

  let updated = 0;
  forEachSelectedDataRow_(sel, function (rowNumber) {
    sh.getRange(rowNumber, COL.STATUS).setValue(newStatus);
    updated++;
  });

  alert_(messageLabel + ': ' + updated + '件');
}

// ============================================================================
// 6. 会員名簿の書き出し
// ============================================================================

/**
 * 入金確認済みの人だけを会員名簿シートに書き出します。
 */
function exportMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  if (!src || src.getLastRow() < 2) {
    alert_('申込・名簿にデータがありません。');
    return;
  }

  const data = src.getRange(2, 1, src.getLastRow() - 1, ROSTER_COLS).getValues();
  const out = [[
    '受付番号',
    '保護者氏名',
    '児童・生徒氏名',
    '学年・組',
    'メールアドレス',
    '会員成立日',
  ]];

  data.forEach(function (r) {
    if (isMember_(r[COL.STATUS - 1])) {
      out.push([
        r[COL.APPLICATION_ID - 1],
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
  applyMemberFormatting_(sh);

  alert_('会員名簿を更新しました。会員数: ' + (out.length - 1) + '件');
}

// ============================================================================
// 7. 支払い案内メール
// ============================================================================

/**
 * 「未納者への督促」ではなく、申込受付中の人への支払い方法の案内として扱います。
 */
function sendPaymentGuides() {
  if (!CONFIG.EMAIL_ENABLED) {
    alert_('メール送信は無効です。テスト完了後に CONFIG.EMAIL_ENABLED を true にしてください。');
    return;
  }

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sh || sh.getLastRow() < 2) {
    alert_('申込・名簿にデータがありません。');
    return;
  }

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, ROSTER_COLS).getValues();
  let sent = 0;
  let skipped = 0;
  let failed = 0;

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const email = String(r[COL.EMAIL - 1] || '').trim();
    if (isPending_(r[COL.STATUS - 1]) && email) {
      const ok = sendPaymentGuideMail_(email, r[COL.GUARDIAN_NAME - 1], r[COL.APPLICATION_ID - 1]);
      const rowNumber = i + 2;
      if (ok) {
        sh.getRange(rowNumber, COL.GUIDE_COUNT).setValue(Number(r[COL.GUIDE_COUNT - 1] || 0) + 1);
        sh.getRange(rowNumber, COL.GUIDE_DATE).setValue(new Date());
        sent++;
        Utilities.sleep(200);
      } else {
        failed++;
      }
    } else {
      skipped++;
    }
  }

  alert_('支払い案内メール送信: ' + sent + '件\n送信失敗: ' + failed + '件\n対象外: ' + skipped + '件');
}

// 旧版のトリガー名を残している場合の互換用。
function sendReminders() {
  sendPaymentGuides();
}

// ============================================================================
// 8. 集計・設定確認
// ============================================================================

function showSummary() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sh || sh.getLastRow() < 2) {
    alert_('申込・名簿にデータがありません。');
    return;
  }

  const statuses = sh.getRange(2, COL.STATUS, sh.getLastRow() - 1, 1).getValues();
  let pending = 0;
  let member = 0;
  let notApplied = 0;
  let withdrawn = 0;
  let review = 0;
  let other = 0;

  statuses.forEach(function (row) {
    const status = String(row[0] || '').trim();
    if (isPending_(status)) pending++;
    else if (isMember_(status)) member++;
    else if (status === STATUS_NOT_APPLIED) notApplied++;
    else if (status === STATUS_WITHDRAWN) withdrawn++;
    else if (status === STATUS_NEEDS_REVIEW) review++;
    else other++;
  });

  alert_(
    '申込受付（入金前）: ' + pending +
    '\n会員（入金確認済）: ' + member +
    '\n申込なし（記録のみ）: ' + notApplied +
    '\n取下げ／無効: ' + withdrawn +
    '\n確認要: ' + review +
    '\nその他: ' + other
  );
}

function showConfig() {
  alert_(
    '団体名: ' + CONFIG.ORG_NAME +
    '\n問い合わせ先: ' + CONFIG.CONTACT_EMAIL +
    '\n同意文の版: ' + CONFIG.CONSENT_VERSION +
    '\n申込管理シート: ' + CONFIG.ROSTER_SHEET +
    '\n会員名簿シート: ' + CONFIG.MEMBER_SHEET +
    '\nメール送信: ' + (CONFIG.EMAIL_ENABLED ? '有効' : '無効') +
    '\n加入申込として扱う選択肢: ' + CONFIG.APPLY_VALUE
  );
}

// ============================================================================
// 9. メール本文
// ============================================================================

function sendReceptionMail_(to, name, id, status) {
  let subject = '【' + CONFIG.ORG_NAME + '】フォーム受付のお知らせ';
  let body = (name || '保護者') + ' 様\n\n';

  if (status === STATUS_APPLIED) {
    subject = '【' + CONFIG.ORG_NAME + '】入会申込受付のお知らせ';
    body +=
      CONFIG.ORG_NAME + 'への入会申込を受け付けました。\n' +
      '受付番号: ' + id + '\n\n' +
      'この時点では「申込受付（入金前）」です。\n' +
      '会費のお支払いを本会が確認した時点で、会員登録が完了します。\n\n' +
      CONFIG.PAYMENT_GUIDE_TEXT + '\n\n' +
      '入会を取りやめる場合や、誤って送信した場合は、下記までご連絡ください。\n\n';
  } else if (status === STATUS_NOT_APPLIED) {
    body +=
      CONFIG.ORG_NAME + 'のフォーム送信を受け付けました。\n' +
      '受付番号: ' + id + '\n\n' +
      '今回の回答は、入会申込としては扱っていません。\n' +
      '内容に誤りがある場合は、下記までご連絡ください。\n\n';
  } else {
    body +=
      CONFIG.ORG_NAME + 'のフォーム送信を受け付けましたが、重要事項への同意等について確認が必要です。\n' +
      '受付番号: ' + id + '\n\n' +
      '必要に応じて、本会からあらためてご連絡します。\n\n';
  }

  body += CONFIG.ORG_NAME + '\n' + CONFIG.CONTACT_EMAIL;
  return safeSendEmail_(to, subject, body);
}

function sendMemberConfirmedMail_(to, name, id) {
  const body =
    (name || '保護者') + ' 様\n\n' +
    CONFIG.ORG_NAME + 'です。\n' +
    '会費のお支払いを確認し、会員登録が完了しました。\n' +
    '受付番号: ' + id + '\n\n' +
    '今後の連絡方法や活動内容については、本会からの案内をご確認ください。\n\n' +
    CONFIG.ORG_NAME + '\n' + CONFIG.CONTACT_EMAIL;

  return safeSendEmail_(to, '【' + CONFIG.ORG_NAME + '】会員登録完了のお知らせ', body);
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
    '※このメールは申込受付中の方へ送信しています。行き違いの場合はご容赦ください。\n\n' +
    CONFIG.ORG_NAME;

  return safeSendEmail_(to, '【' + CONFIG.ORG_NAME + '】会費のお支払い方法のご案内', body);
}

function safeSendEmail_(to, subject, body) {
  try {
    MailApp.sendEmail(to, subject, body);
    return true;
  } catch (err) {
    Logger.log('Mail send failed: ' + err);
    return false;
  }
}

// ============================================================================
// 10. シート作成・補助関数
// ============================================================================

function getOrCreateRoster_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.ROSTER_SHEET);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      '受付番号',
      '受付日時',
      '保護者氏名',
      '児童・生徒氏名',
      '学年・組',
      'メールアドレス',
      '加入意思',
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
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(['受付番号', '保護者氏名', '児童・生徒氏名', '学年・組', 'メールアドレス', '会員成立日']);
  }
  return sh;
}

function getOrCreateGuideSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '使い方';
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 13, 2).setValues([
      ['最初に行うこと', 'Apps Scriptで setupPtaTool を実行し、onFormSubmit トリガーを設定します。'],
      ['フォーム項目', '保護者氏名、児童・生徒氏名、学年・組、メールアドレス、加入意思の確認、重要事項に同意します。'],
      ['質問文の一致', 'フォームの質問文と Code.gs の FIELD_MAP を一字一句そろえてください。'],
      ['申込受付', 'フォーム送信後は、原則として「申込受付（入金前）」です。'],
      ['会員確定', '会費の入金確認後、対象行を選び「会員として確定」を実行します。'],
      ['会員名簿', '「会員名簿を書き出す」を実行すると、入金確認済みの人だけ出力されます。'],
      ['取下げ等', '取下げ、誤送信、重複申込は「取下げ／無効」に変更します。'],
      ['メール送信', 'テスト完了後に CONFIG.EMAIL_ENABLED を true にすると使えます。'],
      ['学校名簿', 'このツールは学校名簿の取り込みを前提にしていません。'],
      ['会費処理', '決済代行ではありません。入金確認は会計担当者が別途行います。'],
      ['同意文の版', '説明文を改訂したら CONFIG.CONSENT_VERSION を更新してください。'],
      ['個人情報', '取得項目は入会管理に必要な範囲に限定してください。'],
      ['引継ぎ', 'PTA専用Googleアカウントで管理し、役員交代時に引き継いでください。'],
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function applyRosterFormatting_(sh) {
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, ROSTER_COLS);

  const maxRows = Math.max(sh.getMaxRows() - 1, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS_APPLIED, STATUS_MEMBER, STATUS_NOT_APPLIED, STATUS_WITHDRAWN, STATUS_NEEDS_REVIEW], true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, COL.STATUS, maxRows, 1).setDataValidation(statusRule);
}

function applyMemberFormatting_(sh) {
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 6);
}

function getRosterActiveSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sh.getName() !== CONFIG.ROSTER_SHEET) {
    alert_('「' + CONFIG.ROSTER_SHEET + '」シートで対象行を選択してから実行してください。');
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

function nextApplicationId_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getDocumentProperties();
    const n = Number(props.getProperty('APPLICATION_SEQ') || '0') + 1;
    props.setProperty('APPLICATION_SEQ', String(n));
    return 'PTA-' + Utilities.formatString('%05d', n);
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

function alert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (err) {
    Logger.log(msg);
  }
}
