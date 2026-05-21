/**
 * portfolio-form-line-notifier
 * Googleフォーム送信 → GAS → LINE Bot 即時通知システム
 *
 * 概要:
 *   Googleフォームに回答が送信されると、指定のLINEグループ/ユーザーへ
 *   即座に通知を届けるシステム。問い合わせ対応・予約受付・アンケート収集など
 *   あらゆる入力フォームに適用可能。
 *
 * セットアップ:
 *   1. GASエディタでこのスクリプトを開く
 *   2. GASのスクリプトプロパティに以下を設定:
 *      - LINE_CHANNEL_ACCESS_TOKEN : LINE Messaging APIのチャンネルアクセストークン
 *      - LINE_TARGET_ID            : 送信先のユーザーID または グループID
 *   3. Googleフォームと紐付いたスプレッドシートのコンテナとしてスクリプトを配置
 *   4. onFormSubmit を「フォーム送信時」トリガーに設定
 *
 * 作成者: hariku0212
 */

// ==============================
// 設定取得（PropertiesServiceで管理）
// ==============================

/**
 * スクリプトプロパティからLINE設定を取得する
 * @returns {{token: string, targetId: string}}
 */
function getLineConfig() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const targetId = props.getProperty('LINE_TARGET_ID');

  if (!token || !targetId) {
    throw new Error(
      'スクリプトプロパティに LINE_CHANNEL_ACCESS_TOKEN と LINE_TARGET_ID を設定してください。'
    );
  }
  return { token, targetId };
}

// ==============================
// フォーム送信トリガー
// ==============================

/**
 * Googleフォーム送信時に呼ばれるメインハンドラ
 * トリガー: フォーム送信時（onFormSubmit）
 *
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e - フォーム送信イベント
 */
function onFormSubmit(e) {
  try {
    const namedValues = e.namedValues;
    const timestamp = e.values[0];

    const message = buildLineMessage(namedValues, timestamp);
    sendLineMessage(message);

    console.log('LINE通知送信完了: ' + timestamp);
  } catch (err) {
    console.error('エラーが発生しました: ' + err.message);
    notifyAdminOnError(err);
  }
}

// ==============================
// メッセージ整形
// ==============================

/**
 * フォーム回答をLINEメッセージテキストに整形する
 *
 * @param {Object} namedValues - { 質問文: [回答] }
 * @param {string} timestamp   - 送信日時文字列
 * @returns {string} LINEに送る整形済みテキスト
 */
function buildLineMessage(namedValues, timestamp) {
  const lines = [];
  lines.push('📬 フォームに新しい回答が届きました');
  lines.push('━━━━━━━━━━━━━━━━');
  lines.push('📅 送信日時: ' + timestamp);
  lines.push('');

  for (const [question, answers] of Object.entries(namedValues)) {
    if (question === 'タイムスタンプ' || question === 'Timestamp') continue;

    const answer = answers.join(', ') || '（未回答）';
    lines.push('▶ ' + question);
    lines.push('   ' + answer);
  }

  lines.push('━━━━━━━━━━━━━━━━');
  lines.push('📊 スプレッドシートで確認: ' + getSpreadsheetUrl());

  return lines.join('\n');
}

function getSpreadsheetUrl() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getUrl();
  } catch (_) {
    return '（URL取得不可）';
  }
}

// ==============================
// LINE Messaging API 送信
// ==============================

/**
 * LINE Messaging API を使ってプッシュメッセージを送信する
 *
 * @param {string} text - 送信するテキスト
 */
function sendLineMessage(text) {
  const config = getLineConfig();

  const payload = {
    to: config.targetId,
    messages: [
      {
        type: 'text',
        text: text
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + config.token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(
    'https://api.line.me/v2/bot/message/push',
    options
  );

  const statusCode = response.getResponseCode();
  if (statusCode !== 200) {
    const body = response.getContentText();
    throw new Error('LINE API エラー (HTTP ' + statusCode + '): ' + body);
  }
}

// ==============================
// エラー通知
// ==============================

function notifyAdminOnError(err) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    if (!adminEmail) return;

    MailApp.sendEmail({
      to: adminEmail,
      subject: '[form-line-notifier] エラーが発生しました',
      body: [
        'フォーム通知システムでエラーが発生しました。',
        '',
        'エラー内容:',
        err.message,
        '',
        'スタックトレース:',
        err.stack || '（取得不可）'
      ].join('\n')
    });
  } catch (_) {
    // メール送信失敗は無視
  }
}

// ==============================
// テスト関数
// ==============================

/**
 * テスト送信（手動実行用）
 */
function testSendLine() {
  const testMessage = [
    '🧪 テスト送信',
    '━━━━━━━━━━━━━━━━',
    '送信日時: ' + new Date().toLocaleString('ja-JP'),
    '',
    '▶ お名前',
    '   テスト 太郎',
    '',
    '▶ お問い合わせ内容',
    '   GASによる業務自動化について相談したい',
    '━━━━━━━━━━━━━━━━',
    'セットアップが正常に完了しています！'
  ].join('\n');

  sendLineMessage(testMessage);
  console.log('テスト送信完了');
}

/**
 * トリガーを自動設定するセットアップ関数（初回のみ実行）
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  console.log('トリガーを設定しました');
}
