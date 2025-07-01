/**
 * Calendar Change Notification Script
 * 
 * 機能:
 * - Google Calendarの定期チェックによる休暇連絡集計
 * - 申請・取消の検知
 * - 対象者ごとの集計報告メール送信
 * - カスタムFromアドレスでのメール送信
 * 
 * @author shiger
 * @version 1.0.0
 * @since 2025-06-27
 */

// ========================================
// 設定値
// ========================================

const CONFIG = {
  // カレンダー設定
  CALENDAR: {
    ID: 'c_651a2cb6c97021756174ac59ac37c04422795de96b016d332847595a35a15ce7@group.calendar.google.com',
    TIMEZONE: 'Asia/Tokyo',
    SEARCH_DAYS_FORWARD: 30,  // 検索期間（日）
    DEFAULT_SEARCH_DAYS_BACK: 1  // デフォルト検索期間（日）
  },
  
  // 通知設定
  NOTIFICATION: {
    HR_EMAIL: 'report@tafs.co.jp',
    ADMIN_EMAIL: 'shigeru.ishizaka@tafs.co.jp',
    SUBJECT_PREFIX: '【勤怠連絡】',
    
    // メール送信オプション
    EMAIL_OPTIONS: {
      from: 'cireport@tafs.co.jp',        // 実行オーナーのエイリアス
      name: 'CI部-勤怠連絡',         // 送信者名
      // replyTo: 'hr@tafs.co.jp',     // 返信先アドレス（必要に応じて有効化）
      // noReply: true                 // 返信不可設定（必要に応じて有効化）
    }
  },
  
  // 対象種別
  TARGET_TYPES: ['全休', '午前休', '午後休', '遅刻', '早退', '特別休', '休出'],
  
  // プロパティキー
  PROPERTY_KEYS: {
    LAST_EXECUTION: 'last_execution_time',
    EVENT_STATE: 'previous_event_state'
  }
};

// メールテンプレート
const EMAIL_TEMPLATE = {
  SUBJECT: '【勤怠連絡】CI部_{REPORT_DATE}',
  BODY: `CI部の勤怠連絡です。

【申請】
{APPLICATIONS}

【取消】
{CANCELLATIONS}

---
このメールは「CI-休暇管理カレンダー」の自動通知メールです。

カレンダー: https://calendar.google.com/calendar/embed?src=${CONFIG.CALENDAR.ID}`.trim()
};

// ========================================
// メイン処理
// ========================================

/**
 * 定期実行メイン関数（トリガーで実行）
 */
function executeAttendanceReport() {
  try {
    const lastExecutionTime = getLastExecutionTime();
    const currentTime = new Date();
    
    const changes = detectChanges(lastExecutionTime, currentTime);
    
    if (changes.hasChanges) {
      sendAttendanceReport(changes, currentTime);
    }
    
    updateLastExecutionTime(currentTime);
    
  } catch (error) {
    console.error('休暇連絡集計報告エラー:', error);
    sendErrorNotification('ATTENDANCE_REPORT', error);
  }
}

/**
 * 変更を検知
 * @param {Date} lastExecutionTime - 前回実行時刻
 * @param {Date} currentTime - 現在時刻
 * @return {object} 変更情報
 */
function detectChanges(lastExecutionTime, currentTime) {
  const calendar = getCalendar();
  const searchPeriod = calculateSearchPeriod(lastExecutionTime, currentTime);
  const events = getAttendanceEvents(calendar, searchPeriod);
  
  const previousState = getPreviousEventState();
  const currentState = createEventState(events);
  const changes = calculateChanges(previousState, currentState);
  
  saveEventState(currentState);
  
  return changes;
}

/**
 * カレンダーを取得
 * @return {Calendar} カレンダーオブジェクト
 * @throws {Error} カレンダーが見つからない場合
 */
function getCalendar() {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR.ID);
  if (!calendar) {
    throw new Error(`カレンダーが見つかりません: ${CONFIG.CALENDAR.ID}`);
  }
  return calendar;
}

/**
 * 検索期間を計算
 * @param {Date} lastExecutionTime - 前回実行時刻
 * @param {Date} currentTime - 現在時刻
 * @return {object} 検索期間
 */
function calculateSearchPeriod(lastExecutionTime, currentTime) {
  const startTime = lastExecutionTime || 
    new Date(currentTime.getTime() - CONFIG.CALENDAR.DEFAULT_SEARCH_DAYS_BACK * 24 * 60 * 60 * 1000);
  
  const endTime = new Date(currentTime.getTime() + CONFIG.CALENDAR.SEARCH_DAYS_FORWARD * 24 * 60 * 60 * 1000);
  
  return { startTime, endTime };
}

/**
 * 勤怠連絡イベントを取得
 * @param {Calendar} calendar - カレンダーオブジェクト
 * @param {object} searchPeriod - 検索期間
 * @return {Array} 休暇連絡イベント配列
 */
function getAttendanceEvents(calendar, searchPeriod) {
  const events = calendar.getEvents(searchPeriod.startTime, searchPeriod.endTime);
  
  return events.filter(event => isAttendanceEvent({
    summary: event.getTitle(),
    description: event.getDescription()
  }));
}

/**
 * 休暇連絡イベントかチェック
 * @param {object} event - カレンダーイベント
 * @return {boolean} 休暇連絡イベントかどうか
 */
function isAttendanceEvent(event) {
  if (!event?.summary) {
    return false;
  }
  
  return CONFIG.TARGET_TYPES.some(type => event.summary.includes(type));
}

/**
 * イベント状態を作成
 * @param {Array} events - イベント配列
 * @return {object} イベント状態
 */
function createEventState(events) {
  const state = {};
  
  events.forEach(event => {
    const eventInfo = extractEventInfo(event);
    state[event.getId()] = eventInfo;
  });
  
  return state;
}

/**
 * イベント情報を抽出
 * @param {object} event - カレンダーイベント
 * @return {object} イベント情報
 */
function extractEventInfo(event) {
  const summary = event.getTitle();
  const description = event.getDescription();
  
  const { personName, attendanceType } = parseEventSummary(summary);
  const reporter = extractReporter(description);
  
  return {
    id: event.getId(),
    summary,
    description,
    personName,
    attendanceType,
    reporter,
    startTime: event.getStartTime(),
    endTime: event.getEndTime(),
    isAllDay: event.isAllDayEvent(),
    lastUpdated: event.getLastUpdated()
  };
}

/**
 * イベントサマリーを解析
 * @param {string} summary - イベントサマリー
 * @return {object} 解析結果
 */
function parseEventSummary(summary) {
  const nameMatch = summary.match(/^(.+?)\s*-\s*(.+)$/);
  
  return {
    personName: nameMatch ? nameMatch[1].trim() : '不明',
    attendanceType: nameMatch ? nameMatch[2].trim() : '不明'
  };
}

/**
 * 申請者を抽出
 * @param {string} description - イベント説明
 * @return {string} 申請者名
 */
function extractReporter(description) {
  const reporterMatch = description.match(/申請者:\s*(.+?)(?:\n|$)/);
  return reporterMatch ? reporterMatch[1].trim() : '不明';
}

/**
 * 変更を計算
 * @param {object} previousState - 前回の状態
 * @param {object} currentState - 現在の状態
 * @return {object} 変更情報
 */
function calculateChanges(previousState, currentState) {
  const changes = {
    hasChanges: false,
    applications: [],    // 申請（新規）
    cancellations: []    // 取消（削除）
  };
  
  // 申請（新規登録）を検出
  Object.keys(currentState).forEach(eventId => {
    if (!previousState[eventId]) {
      changes.applications.push(currentState[eventId]);
      changes.hasChanges = true;
    }
  });
  
  // 取消（削除）を検出
  Object.keys(previousState).forEach(eventId => {
    if (!currentState[eventId]) {
      changes.cancellations.push(previousState[eventId]);
      changes.hasChanges = true;
    }
  });
  
  return changes;
}

/**
 * 休暇連絡報告を送信
 * @param {object} changes - 変更情報
 * @param {Date} currentTime - 現在時刻
 */
function sendAttendanceReport(changes, currentTime) {
  const reportDate = formatReportDate(currentTime);
  const personChanges = aggregateChangesByPerson(changes);
  
  const subject = EMAIL_TEMPLATE.SUBJECT.replace('{REPORT_DATE}', reportDate);
  const body = createEmailBody(personChanges);
  
  sendEmail(CONFIG.NOTIFICATION.HR_EMAIL, subject, body);
}

/**
 * 報告日時をフォーマット
 * @param {Date} date - 日付
 * @return {string} フォーマット済み日時
 */
function formatReportDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hour = String(date.getHours()).padStart(2, '0');
  
  return `${year}/${month}/${day}-${hour}時`;
}

/**
 * 変更を対象者ごとに集計
 * @param {object} changes - 変更情報
 * @return {object} 対象者別集計
 */
function aggregateChangesByPerson(changes) {
  const personChanges = {};
  
  // 申請を集計
  changes.applications.forEach(event => {
    addPersonChange(personChanges, event.personName, 'applications', event);
  });
  
  // 取消を集計
  changes.cancellations.forEach(event => {
    addPersonChange(personChanges, event.personName, 'cancellations', event);
  });
  
  return personChanges;
}

/**
 * 対象者別変更を追加
 * @param {object} personChanges - 対象者別変更
 * @param {string} personName - 対象者名
 * @param {string} changeType - 変更種別
 * @param {object} event - イベント情報
 */
function addPersonChange(personChanges, personName, changeType, event) {
  if (!personChanges[personName]) {
    personChanges[personName] = { applications: [], cancellations: [] };
  }
  
  personChanges[personName][changeType].push({
    type: event.attendanceType,
    date: event.startTime.toLocaleDateString('ja-JP')
  });
}

/**
 * メール本文を作成
 * @param {object} personChanges - 対象者別変更
 * @return {string} メール本文
 */
function createEmailBody(personChanges) {
  const applicationsText = formatChangeSection(personChanges, 'applications', '申請');
  const cancellationsText = formatChangeSection(personChanges, 'cancellations', '取消');
  
  return EMAIL_TEMPLATE.BODY
    .replace('{APPLICATIONS}', applicationsText)
    .replace('{CANCELLATIONS}', cancellationsText);
}

/**
 * 変更セクションをフォーマット
 * @param {object} personChanges - 対象者別変更
 * @param {string} changeType - 変更種別
 * @param {string} sectionName - セクション名
 * @return {string} フォーマット済みテキスト
 */
function formatChangeSection(personChanges, changeType, sectionName) {
  const hasChanges = Object.values(personChanges).some(changes => changes[changeType].length > 0);
  
  if (!hasChanges) {
    return '';
  }
  
  const formattedEntries = [];
  
  Object.entries(personChanges)
    .filter(([person, changes]) => changes[changeType].length > 0)
    .forEach(([person, changes]) => {
      changes[changeType].forEach(event => {
        formattedEntries.push(`    ${person}: ${event.type} (${event.date})`);
      });
    });
  
  return formattedEntries.join('\n');
}

/**
 * メール送信
 * @param {string} to - 送信先
 * @param {string} subject - 件名
 * @param {string} body - 本文
 */
function sendEmail(to, subject, body) {
  try {
    const emailOptions = {
      ...CONFIG.NOTIFICATION.EMAIL_OPTIONS
    };
    
    GmailApp.sendEmail(to, subject, body, emailOptions);
  } catch (error) {
    console.error('メール送信エラー:', error);
    throw error;
  }
}

// ========================================
// 状態管理関数
// ========================================

/**
 * 前回実行時刻を取得
 * @return {Date} 前回実行時刻
 */
function getLastExecutionTime() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const lastExecutionStr = properties.getProperty(CONFIG.PROPERTY_KEYS.LAST_EXECUTION);
    
    if (lastExecutionStr) {
      return new Date(lastExecutionStr);
    }
    
    return new Date(Date.now() - CONFIG.CALENDAR.DEFAULT_SEARCH_DAYS_BACK * 24 * 60 * 60 * 1000);
    
  } catch (error) {
    console.error('前回実行時刻取得エラー:', error);
    return new Date(Date.now() - CONFIG.CALENDAR.DEFAULT_SEARCH_DAYS_BACK * 24 * 60 * 60 * 1000);
  }
}

/**
 * 実行時刻を更新
 * @param {Date} executionTime - 実行時刻
 */
function updateLastExecutionTime(executionTime) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(CONFIG.PROPERTY_KEYS.LAST_EXECUTION, executionTime.toISOString());
  } catch (error) {
    console.error('実行時刻更新エラー:', error);
  }
}

/**
 * 前回のイベント状態を取得
 * @return {object} イベント状態
 */
function getPreviousEventState() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const stateJson = properties.getProperty(CONFIG.PROPERTY_KEYS.EVENT_STATE);
    
    if (!stateJson) {
      return {};
    }
    
    const state = JSON.parse(stateJson);
    return convertDateStringsToDates(state);
    
  } catch (error) {
    console.error('前回イベント状態取得エラー:', error);
    return {};
  }
}

/**
 * 日付文字列をDateオブジェクトに変換
 * @param {object} state - イベント状態
 * @return {object} 変換済みイベント状態
 */
function convertDateStringsToDates(state) {
  Object.keys(state).forEach(eventId => {
    const event = state[eventId];
    
    if (event.startTime) {
      event.startTime = new Date(event.startTime);
    }
    if (event.endTime) {
      event.endTime = new Date(event.endTime);
    }
    if (event.lastUpdated) {
      event.lastUpdated = new Date(event.lastUpdated);
    }
  });
  
  return state;
}

/**
 * イベント状態を保存
 * @param {object} state - イベント状態
 */
function saveEventState(state) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const stateForSave = convertDatesToStrings(state);
    
    properties.setProperty(CONFIG.PROPERTY_KEYS.EVENT_STATE, JSON.stringify(stateForSave));
    
  } catch (error) {
    console.error('イベント状態保存エラー:', error);
  }
}

/**
 * Dateオブジェクトを文字列に変換
 * @param {object} state - イベント状態
 * @return {object} 変換済みイベント状態
 */
function convertDatesToStrings(state) {
  const stateForSave = {};
  
  Object.keys(state).forEach(eventId => {
    const event = state[eventId];
    
    stateForSave[eventId] = {
      ...event,
      startTime: event.startTime ? event.startTime.toISOString() : null,
      endTime: event.endTime ? event.endTime.toISOString() : null,
      lastUpdated: event.lastUpdated ? event.lastUpdated.toISOString() : null
    };
  });
  
  return stateForSave;
}

// ========================================
// エラー処理
// ========================================

/**
 * エラー通知メールを送信
 * @param {string} type - 通知タイプ
 * @param {Error} error - エラーオブジェクト
 */
function sendErrorNotification(type, error) {
  const subject = `${CONFIG.NOTIFICATION.SUBJECT_PREFIX} エラー通知 - ${type}`;
  const body = createErrorEmailBody(type, error);
  
  try {
    sendEmail(CONFIG.NOTIFICATION.ADMIN_EMAIL, subject, body);
  } catch (emailError) {
    console.error('エラー通知メール送信エラー:', emailError);
  }
}

/**
 * エラーメール本文を作成
 * @param {string} type - 通知タイプ
 * @param {Error} error - エラーオブジェクト
 * @return {string} メール本文
 */
function createErrorEmailBody(type, error) {
  return `休暇管理システムでエラーが発生しました。

【エラー詳細】
タイプ: ${type}
エラー: ${error.message}
スタックトレース: ${error.stack}
発生時刻: ${new Date().toLocaleString('ja-JP', { timeZone: CONFIG.CALENDAR.TIMEZONE })}

---
このメールは休暇管理システムにより自動送信されています。`.trim();
}

// ========================================
// テスト・設定用関数
// ========================================

/**
 * 手動実行テスト
 */
function testAttendanceReport() {
  executeAttendanceReport();
}

/**
 * 設定確認
 */
function checkConfiguration() {
  console.log('=== 設定確認 ===');
  console.log('カレンダーID:', CONFIG.CALENDAR.ID);
  console.log('タイムゾーン:', CONFIG.CALENDAR.TIMEZONE);
  console.log('対象種別:', CONFIG.TARGET_TYPES);
  console.log('人事メール:', CONFIG.NOTIFICATION.HR_EMAIL);
  console.log('管理者メール:', CONFIG.NOTIFICATION.ADMIN_EMAIL);
  console.log('検索期間（日）:', CONFIG.CALENDAR.SEARCH_DAYS_FORWARD);
  console.log('デフォルト検索期間（日）:', CONFIG.CALENDAR.DEFAULT_SEARCH_DAYS_BACK);
  console.log('送信者アドレス:', CONFIG.NOTIFICATION.EMAIL_OPTIONS.from);
  console.log('送信者名:', CONFIG.NOTIFICATION.EMAIL_OPTIONS.name);
}

/**
 * トリガー設定
 */
function setupTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'executeAttendanceReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成（毎日9:00と17:00に実行）
  ScriptApp.newTrigger('executeAttendanceReport')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  ScriptApp.newTrigger('executeAttendanceReport')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .create();
}

/**
 * 保存データをクリア
 */
function clearStoredData() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty(CONFIG.PROPERTY_KEYS.LAST_EXECUTION);
    properties.deleteProperty(CONFIG.PROPERTY_KEYS.EVENT_STATE);
    console.log('保存データをクリアしました');
  } catch (error) {
    console.error('データクリアエラー:', error);
  }
} 