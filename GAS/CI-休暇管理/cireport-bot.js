/**
 * Attendance Management Chatbot for Google Chat
 * 
 * 機能:
 * - 勤怠連絡の受付とGoogle Calendarへの自動登録
 * - 複数日付・範囲指定対応
 * - 勤怠取消機能
 * - 申請者権限による削除制限
 * - デバッグログ機能
 * 
 * 対応種別: 全休、午前休、午後休、遅刻、早退、特別休、休出、取消
 * 
 * @author shiger
 * @version 1.0.0
 * @since 2025-06-27
 */

// ========================================
// 設定値
// ========================================

// 時間定数（ミリ秒）
const TIME_CONSTANTS = {
  HOUR: 60 * 60 * 1000,
  MINUTE: 60 * 1000,
  WORK_START: 9 * 60 * 60 * 1000,    // 9:00
  WORK_END: 17.5 * 60 * 60 * 1000,   // 17:30
  LUNCH_START: 12 * 60 * 60 * 1000,  // 12:00
  LUNCH_END: 13 * 60 * 60 * 1000,    // 13:00
  LATE_START: 9 * 60 * 60 * 1000,    // 9:00
  LATE_END: 10.5 * 60 * 60 * 1000,   // 10:30
  EARLY_LEAVE_START: 16 * 60 * 60 * 1000, // 16:00
  EARLY_LEAVE_END: 17.5 * 60 * 60 * 1000  // 17:30
};

// Bot設定
const BOT_CONFIG = {
  CALENDAR_ID: 'c_651a2cb6c97021756174ac59ac37c04422795de96b016d332847595a35a15ce7@group.calendar.google.com',
  TIMEZONE: 'Asia/Tokyo'
};

// 正規表現パターン
const REGEX_PATTERNS = {
  DATE_FORMAT: /^\d{8}$/,
  ATTENDANCE_MESSAGE: /【勤怠連絡】\s*\n\s*氏名：\s*([^\s]+(?:\s+[^\s]+)*)\s*\n\s*種別：\s*(全休|午前休|午後休|遅刻|早退|特別休|休出|取消)\s*\n\s*日付：\s*(.+)\s*\n\s*備考：\s*(.+)/s,
  MENTION: /@[\w-]+/g
};

// メッセージ設定
const MESSAGE_CONFIG = {
  // ヘルプメッセージ
  HELP_MESSAGE: {
    GREETING: 'こんにちは、{userName}さん！\n勤怠連絡を受け付けています。\n\n',
    FORMAT_TITLE: '【勤怠連絡フォーマット】\n',
    FORMAT_HEADER: '【勤怠連絡】\n',
    FORMAT_NAME: '氏名：　〇〇　〇〇\n',
    FORMAT_TYPE: '種別：　[全休/午前休/午後休/遅刻/早退/特別休/休出/取消] から選択\n',
    FORMAT_DATE: '日付：　YYYYMMDD (複数日付: カンマ区切り、範囲指定: YYYYMMDD-YYYYMMDD)\n',
    FORMAT_REMARKS: '備考：　〇〇のため\n\n',
    DATE_EXAMPLES: '【日付指定例】\n・単一日付: 20250115\n・複数日付: 20250115,20250116,20250117\n・範囲指定: 20250115-20250117\n\n',
    CANCELLATION_INFO: '【取消について】\n種別に「取消」を指定すると、指定日付の予定を削除します。'
  },
  
  // エラーメッセージ
  ERROR_MESSAGE: {
    FORMAT_ERROR: '申し訳ございません。フォーマットが正しくありません。\n\n',
    CALENDAR_ERROR: '❌ 申し訳ございません。カレンダーへの追加に失敗しました。\nエラー: ',
    DELETE_ERROR: '❌ 申し訳ございません。イベントの削除に失敗しました。\nエラー: '
  },
  
  // 成功メッセージ
  SUCCESS_MESSAGE: {
    ATTENDANCE_RECEIVED: '✅ 勤怠連絡を受け付けました！\n',
    CANCELLATION_RECEIVED: '✅ 勤怠取消を受け付けました！\n',
    EVENTS_ADDED: '✅ {count}件の予定をカレンダーに追加しました。',
    EVENTS_DELETED: '✅ {personName}さんの{count}件の予定をカレンダーから削除しました。',
    NO_EVENTS_TO_DELETE: 'ℹ️ 指定された日付に{personName}さんの削除対象の予定はありませんでした。',
    ADD_FAILED: '\n❌ {count}件の追加に失敗しました。',
    DELETE_FAILED: '\n❌ {count}件の削除に失敗しました。'
  },
  
  // 表示項目
  DISPLAY_FIELDS: {
    REPORTER: '申請: ',
    NAME: '\n氏名: ',
    TYPE: '種別: ',
    DATE: '日付: ',
    REMARKS: '備考: ',
    SEPARATOR: '─────────────────\n'
  }
};

// 勤怠種別設定
const ATTENDANCE_TYPES = {
  '全休': {
    title: '全休',
    colorId: '4', // 赤色
    startTime: null,
    endTime: null
  },
  '午前休': {
    title: '午前休',
    colorId: '6', // オレンジ色
    startTime: TIME_CONSTANTS.WORK_START,
    endTime: TIME_CONSTANTS.LUNCH_START
  },
  '午後休': {
    title: '午後休',
    colorId: '6', // オレンジ色
    startTime: TIME_CONSTANTS.LUNCH_END,
    endTime: TIME_CONSTANTS.WORK_END
  },
  '遅刻': {
    title: '遅刻',
    colorId: '5', // 黄色
    startTime: TIME_CONSTANTS.LATE_START,
    endTime: TIME_CONSTANTS.LATE_END
  },
  '早退': {
    title: '早退',
    colorId: '5', // 黄色
    startTime: TIME_CONSTANTS.EARLY_LEAVE_START,
    endTime: TIME_CONSTANTS.EARLY_LEAVE_END
  },
  '特別休': {
    title: '特別休',
    colorId: '3', // 緑色
    startTime: null,
    endTime: null
  },
  '休出': {
    title: '休出',
    colorId: '2', // 青色
    startTime: null,
    endTime: null
  },
  '取消': {
    title: '取消',
    colorId: '1', // デフォルト色
    startTime: null,
    endTime: null
  }
};

// 定数
const CONSTANTS = {
  MAX_NAME_LENGTH: 50,
  MAX_REMARKS_LENGTH: 200,
  MAX_DATES_COUNT: 31,
  DEFAULT_COLOR_ID: '1'
};

// ========================================
// ユーティリティ関数
// ========================================

/**
 * 日付文字列をフォーマット（YYYY-MM-DD）
 * @param {Date} date - 日付オブジェクト
 * @return {string} フォーマットされた日付文字列
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 日本時間の日付オブジェクトを作成
 * @param {object} dateData - 日付データ
 * @return {Date} 日本時間の日付オブジェクト
 */
function createJapaneseDate(dateData) {
  const year = dateData.year;
  const month = String(dateData.month).padStart(2, '0');
  const day = String(dateData.day).padStart(2, '0');
  return new Date(`${year}-${month}-${day}T00:00:00+09:00`);
}

/**
 * 日付オブジェクトから文字列を生成
 * @param {object} dateData - 日付データ
 * @return {string} YYYYMMDD形式の文字列
 */
function generateDateString(dateData) {
  return dateData.year.toString() + 
         String(dateData.month).padStart(2, '0') + 
         String(dateData.day).padStart(2, '0');
}

/**
 * 日付範囲を生成する
 * @param {object} startDate - 開始日
 * @param {object} endDate - 終了日
 * @return {Array} 日付オブジェクトの配列
 */
function generateDateRange(startDate, endDate) {
  const dates = [];
  
  // 日付オブジェクトを作成して有効性をチェック
  const start = new Date(startDate.year, startDate.month - 1, startDate.day);
  const end = new Date(endDate.year, endDate.month - 1, endDate.day);
  
  // 日付が有効かチェック
  if (start.getFullYear() !== startDate.year || start.getMonth() !== startDate.month - 1 || start.getDate() !== startDate.day) {
    return dates; // 開始日が無効
  }
  if (end.getFullYear() !== endDate.year || end.getMonth() !== endDate.month - 1 || end.getDate() !== endDate.day) {
    return dates; // 終了日が無効
  }
  
  if (start > end) {
    return dates; // 開始日が終了日より後の場合は空配列を返す
  }
  
  const current = new Date(start);
  while (current <= end) {
    const year = current.getFullYear();
    const month = current.getMonth() + 1;
    const day = current.getDate();
    
    dates.push({
      year,
      month,
      day,
      date: generateDateString({ year, month, day })
    });
    current.setDate(current.getDate() + 1);
  }
  
  return dates;
}

// ========================================
// メイン処理
// ========================================

function onMessage(e) {
  // Bot自身のメッセージには反応しない
  if (e.user.type === 'BOT') {
    return;
  }

  // メンション判定
  const isMentioned = e.message.annotations && e.message.annotations.some(
    annotation => annotation.type === 'USER_MENTION'
  );

  if (!isMentioned) {
    return;
  }

  let userMessage = ''; // デフォルト値を設定
  if (e.message.argumentText !== undefined) {
    userMessage = e.message.argumentText.trim();
  } else {
    // argumentText がない場合の処理（例: 空文字列として扱う）
  }
  
  let replyText;
  let replyCards = [];

  if (!userMessage) {
    replyText = createHelpMessage(e.user.displayName);
  } else {
    const attendanceData = parseAttendanceMessage(userMessage);
    
    if (attendanceData) {
      if (attendanceData.type === '取消') {
        replyText = removeFromCalendar(attendanceData, e.user.displayName);
      } else {
        replyText = addToCalendar(attendanceData, e.user.displayName);
      }
      
      // カレンダーリンクカードを追加（報告対象の日付情報を渡す）
      replyCards = [createCalendarLinkCard(attendanceData.dates)];
    } else {
      replyText = createFormatErrorMessage();
    }
  }

  return { 
    'text': replyText,
    'cards': replyCards
  };
}

/**
 * ヘルプメッセージを作成
 * @param {string} userName - ユーザー名
 * @return {string} ヘルプメッセージ
 */
function createHelpMessage(userName) {
  const helpParts = [
    MESSAGE_CONFIG.HELP_MESSAGE.GREETING.replace('{userName}', userName),
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_TITLE,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_HEADER,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_NAME,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_TYPE,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_DATE,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_REMARKS,
    MESSAGE_CONFIG.HELP_MESSAGE.DATE_EXAMPLES,
    MESSAGE_CONFIG.HELP_MESSAGE.CANCELLATION_INFO
  ];
  
  return helpParts.join('');
}

/**
 * フォーマットエラーメッセージを作成
 * @return {string} エラーメッセージ
 */
function createFormatErrorMessage() {
  const errorParts = [
    MESSAGE_CONFIG.ERROR_MESSAGE.FORMAT_ERROR,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_TITLE,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_HEADER,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_NAME,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_TYPE,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_DATE,
    MESSAGE_CONFIG.HELP_MESSAGE.FORMAT_REMARKS,
    MESSAGE_CONFIG.HELP_MESSAGE.DATE_EXAMPLES,
    MESSAGE_CONFIG.HELP_MESSAGE.CANCELLATION_INFO
  ];
  
  return errorParts.join('');
}

/**
 * 勤怠連絡メッセージを解析する
 * @param {string} message - ユーザーのメッセージ
 * @return {object|null} 解析結果またはnull
 */
function parseAttendanceMessage(message) {
  // メンションを除去してメッセージをクリーンアップ
  const cleanMessage = removeMentions(message);
  
  const match = cleanMessage.match(REGEX_PATTERNS.ATTENDANCE_MESSAGE);
  if (!match) {
    return null;
  }

  const [, name, type, dateInput, remarks] = match;
  
  // 入力値の検証とサニタイゼーション
  const sanitizedName = name.trim();
  const sanitizedRemarks = remarks.trim();
  
  // 文字列長の制限チェック
  if (sanitizedName.length > CONSTANTS.MAX_NAME_LENGTH) {
    return null;
  }
  if (sanitizedRemarks.length > CONSTANTS.MAX_REMARKS_LENGTH) {
    return null;
  }
  
  // 日付の解析（複数日付、範囲指定に対応）
  const dates = parseDateInput(dateInput);
  if (!dates || dates.length === 0) {
    return null;
  }
  
  // 日付数の制限チェック
  if (dates.length > CONSTANTS.MAX_DATES_COUNT) {
    return null;
  }

  return {
    name: sanitizedName,
    type: type,
    dates: dates,
    originalDateInput: dateInput.trim(), // 元の日付入力形式を保持
    remarks: sanitizedRemarks
  };
}

/**
 * メッセージからメンションを除去する
 * @param {string} message - 元のメッセージ
 * @return {string} メンションを除去したメッセージ
 */
function removeMentions(message) {
  return message.replace(REGEX_PATTERNS.MENTION, '').trim();
}

/**
 * 日付入力を解析する（複数日付、範囲指定に対応）
 * @param {string} dateInput - 日付入力文字列
 * @return {Array|null} 日付オブジェクトの配列またはnull
 */
function parseDateInput(dateInput) {
  const dates = [];
  const input = dateInput.trim();
  
  // カンマで区切られた複数日付
  if (input.includes(',')) {
    const dateStrings = input.split(',').map(s => s.trim());
    for (const dateStr of dateStrings) {
      const date = parseSingleDate(dateStr);
      if (date) {
        dates.push(date);
      }
    }
  }
  // 範囲指定（YYYYMMDD-YYYYMMDD）
  else if (input.includes('-')) {
    const range = input.split('-').map(s => s.trim());
    if (range.length === 2) {
      const startDate = parseSingleDate(range[0]);
      const endDate = parseSingleDate(range[1]);
      if (startDate && endDate) {
        const dateRange = generateDateRange(startDate, endDate);
        dates.push(...dateRange);
      }
    }
  }
  // 単一日付
  else {
    const date = parseSingleDate(input);
    if (date) {
      dates.push(date);
    }
  }
  
  return dates.length > 0 ? dates : null;
}

/**
 * 単一日付を解析する
 * @param {string} dateStr - 日付文字列（YYYYMMDD）
 * @return {object|null} 日付オブジェクトまたはnull
 */
function parseSingleDate(dateStr) {
  if (!REGEX_PATTERNS.DATE_FORMAT.test(dateStr)) {
    return null;
  }
  
  const year = parseInt(dateStr.substring(0, 4));
  const month = parseInt(dateStr.substring(4, 6));
  const day = parseInt(dateStr.substring(6, 8));
  
  // 基本的な範囲チェック
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return null;
  }
  
  // 実際の日付として有効かチェック
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null; // 無効な日付（例：2月30日）
  }
  
  return { year, month, day, date: dateStr };
}

/**
 * Google Calendarに予定を追加する
 * @param {object} attendanceData - 勤怠データ
 * @param {string} reporterName - 申請者名
 * @return {string} 結果メッセージ
 */
function addToCalendar(attendanceData, reporterName) {
  try {
    // 種別に応じたイベントタイトルと時間を設定
    const eventConfig = getEventConfig(attendanceData.type);
    
    let successCount = 0;
    let errorCount = 0;
    
    // 各日付に対してイベントを作成
    for (const dateData of attendanceData.dates) {
      try {
        // 日本時間で日付オブジェクトを作成
        const eventDate = createJapaneseDate(dateData);
        
        // イベントオブジェクトを作成
        const event = createEventObject(attendanceData, eventConfig, eventDate, reporterName);
        
        // カレンダーにイベントを追加
        Calendar.Events.insert(event, BOT_CONFIG.CALENDAR_ID);
        successCount++;
        
      } catch (error) {
        errorCount++;
      }
    }
    
    return createSuccessMessage(attendanceData, successCount, errorCount, reporterName);
    
  } catch (error) {
    return MESSAGE_CONFIG.ERROR_MESSAGE.CALENDAR_ERROR + error.toString();
  }
}

/**
 * イベントオブジェクトを作成
 * @param {object} attendanceData - 勤怠データ
 * @param {object} eventConfig - イベント設定
 * @param {Date} eventDate - イベント日付
 * @param {string} reporterName - 申請者名
 * @return {object} イベントオブジェクト
 */
function createEventObject(attendanceData, eventConfig, eventDate, reporterName) {
  const event = {
    'summary': `${attendanceData.name} - ${eventConfig.title}`,
    'description': `備考: ${attendanceData.remarks}\n申請者: ${reporterName}`,
    'colorId': eventConfig.colorId || CONSTANTS.DEFAULT_COLOR_ID
  };

  // 時間指定がある場合とない場合で分岐
  if (eventConfig.startTime !== null && eventConfig.endTime !== null) {
    // 時間指定がある場合（午前休、午後休、遅刻、早退）
    const startDateTime = createTimeSpecificDate(eventDate, eventConfig.startTime);
    const endDateTime = createTimeSpecificDate(eventDate, eventConfig.endTime);
    
    event.start = {
      'dateTime': startDateTime.toISOString(),
      'timeZone': BOT_CONFIG.TIMEZONE
    };
    event.end = {
      'dateTime': endDateTime.toISOString(),
      'timeZone': BOT_CONFIG.TIMEZONE
    };
  } else {
    // 終日イベントの場合（全休、特別休、休出）
    const dateString = formatDate(eventDate);
    const nextDate = new Date(eventDate);
    nextDate.setDate(nextDate.getDate() + 1);
    const nextDateString = formatDate(nextDate);
    
    event.start = {
      'date': dateString,
      'timeZone': BOT_CONFIG.TIMEZONE
    };
    event.end = {
      'date': nextDateString,
      'timeZone': BOT_CONFIG.TIMEZONE
    };
  }

  return event;
}

/**
 * 時間指定の日付オブジェクトを作成
 * @param {Date} baseDate - 基準日付
 * @param {number} timeInMs - 時間（ミリ秒）
 * @return {Date} 時間指定の日付オブジェクト
 */
function createTimeSpecificDate(baseDate, timeInMs) {
  const dateTime = new Date(baseDate);
  const hours = Math.floor(timeInMs / TIME_CONSTANTS.HOUR);
  const minutes = Math.floor((timeInMs % TIME_CONSTANTS.HOUR) / TIME_CONSTANTS.MINUTE);
  
  dateTime.setHours(hours, minutes);
  return dateTime;
}

/**
 * 勤怠情報の表示部分を作成
 * @param {object} attendanceData - 勤怠データ
 * @param {string} reporterName - 申請者名
 * @param {string} dateList - 日付リスト文字列
 * @return {string} 勤怠情報の表示部分
 */
function createAttendanceDisplay(attendanceData, reporterName, dateList) {
  return MESSAGE_CONFIG.DISPLAY_FIELDS.NAME + attendanceData.name + '\n' +
         MESSAGE_CONFIG.DISPLAY_FIELDS.TYPE + attendanceData.type + '\n' +
         MESSAGE_CONFIG.DISPLAY_FIELDS.DATE + dateList + '\n' +
         MESSAGE_CONFIG.DISPLAY_FIELDS.REMARKS + attendanceData.remarks + '\n' +
         MESSAGE_CONFIG.DISPLAY_FIELDS.SEPARATOR +
         MESSAGE_CONFIG.DISPLAY_FIELDS.REPORTER + reporterName + '\n\n';
}

/**
 * 成功メッセージを作成
 * @param {object} attendanceData - 勤怠データ
 * @param {number} successCount - 成功件数
 * @param {number} errorCount - エラー件数
 * @param {string} reporterName - 申請者名
 * @return {string} 成功メッセージ
 */
function createSuccessMessage(attendanceData, successCount, errorCount, reporterName) {
  // 元の日付入力形式を使用
  const dateList = attendanceData.originalDateInput;
  
  let message = MESSAGE_CONFIG.SUCCESS_MESSAGE.ATTENDANCE_RECEIVED +
                createAttendanceDisplay(attendanceData, reporterName, dateList);
  
  if (successCount > 0) {
    message += MESSAGE_CONFIG.SUCCESS_MESSAGE.EVENTS_ADDED.replace('{count}', successCount);
  }
  
  if (errorCount > 0) {
    message += MESSAGE_CONFIG.SUCCESS_MESSAGE.ADD_FAILED.replace('{count}', errorCount);
  }
  
  return message;
}

/**
 * 種別に応じたイベント設定を取得
 * @param {string} type - 勤怠種別
 * @return {object} イベント設定
 */
function getEventConfig(type) {
  return ATTENDANCE_TYPES[type] || ATTENDANCE_TYPES['全休'];
}

/**
 * Google Calendarから予定を削除する
 * @param {object} attendanceData - 勤怠データ
 * @param {string} reporterName - 申請者名
 * @return {string} 結果メッセージ
 */
function removeFromCalendar(attendanceData, reporterName) {
  try {
    let successCount = 0;
    let errorCount = 0;
    const deletedDates = [];
    
    for (const dateData of attendanceData.dates) {
      try {
        const events = searchEventsByDate(BOT_CONFIG.CALENDAR_ID, dateData, attendanceData.name, reporterName, true);
        
        if (events.length > 0) {
          for (const event of events) {
            Calendar.Events.remove(BOT_CONFIG.CALENDAR_ID, event.id);
            successCount++;
          }
          deletedDates.push(`${dateData.year}/${dateData.month}/${dateData.day}`);
        }
        
      } catch (error) {
        errorCount++;
      }
    }
    
    return createCancellationMessage(attendanceData, successCount, errorCount, reporterName, deletedDates);
    
  } catch (error) {
    return MESSAGE_CONFIG.ERROR_MESSAGE.DELETE_ERROR + error.toString();
  }
}

/**
 * 指定日付のイベントを検索する（取消処理用）
 * @param {string} calendarId - カレンダーID
 * @param {object} dateData - 日付データ
 * @param {string} personName - 検索対象の氏名
 * @param {string} reporterName - 申請者名（取消時のみ使用）
 * @param {boolean} isCancellation - 取消処理かどうか
 * @return {Array} イベントの配列
 */
function searchEventsByDate(calendarId, dateData, personName, reporterName = null, isCancellation = false) {
  const startDate = new Date(`${dateData.year}-${String(dateData.month).padStart(2, '0')}-${String(dateData.day).padStart(2, '0')}T00:00:00+09:00`);
  const endDate = new Date(`${dateData.year}-${String(dateData.month).padStart(2, '0')}-${String(dateData.day).padStart(2, '0')}T23:59:59+09:00`);
  
  const events = Calendar.Events.list(calendarId, {
    timeMin: startDate.toISOString(),
    timeMax: endDate.toISOString(),
    singleEvents: true,
    orderBy: 'startTime'
  });
  
  // 氏名でフィルタリング（スペースを除去して比較）
  const normalizedPersonName = personName.replace(/\s+/g, '');
  let filteredEvents = (events.items || []).filter(event => {
    if (!event.summary) return false;
    const normalizedEventName = event.summary.replace(/\s+/g, '');
    return normalizedEventName.includes(normalizedPersonName);
  });
  
  // 取消の場合は申請者名でもフィルタリング
  if (isCancellation && reporterName) {
    filteredEvents = filteredEvents.filter(event => {
      return event.description && event.description.includes(`申請者: ${reporterName}`);
    });
  }
  
  return filteredEvents;
}

/**
 * 取消メッセージを作成
 * @param {object} attendanceData - 勤怠データ
 * @param {number} successCount - 成功件数
 * @param {number} errorCount - エラー件数
 * @param {string} reporterName - 申請者名
 * @param {Array} deletedDates - 実際に削除した日付の配列
 * @return {string} 取消メッセージ
 */
function createCancellationMessage(attendanceData, successCount, errorCount, reporterName, deletedDates) {
  // 申請内容の表示では元の日付形式を使用
  const originalDateList = attendanceData.originalDateInput;
  
  let message = MESSAGE_CONFIG.SUCCESS_MESSAGE.CANCELLATION_RECEIVED +
                createAttendanceDisplay(attendanceData, reporterName, originalDateList);
  
  if (successCount > 0) {
    message += MESSAGE_CONFIG.SUCCESS_MESSAGE.EVENTS_DELETED.replace('{count}', successCount).replace('{personName}', attendanceData.name);
  } else {
    message += MESSAGE_CONFIG.SUCCESS_MESSAGE.NO_EVENTS_TO_DELETE.replace('{personName}', attendanceData.name);
  }
  
  if (errorCount > 0) {
    message += MESSAGE_CONFIG.SUCCESS_MESSAGE.DELETE_FAILED.replace('{count}', errorCount);
  }
  
  return message;
}

/**
 * カレンダーリンクカードを作成
 * @param {Array} dates - 日付オブジェクトの配列
 * @return {object} カレンダーリンクカードオブジェクト
 */
function createCalendarLinkCard(dates) {
  // 日付をソートして最初と最後の日付を取得
  const sortedDates = dates.sort((a, b) => {
    const dateA = new Date(a.year, a.month - 1, a.day);
    const dateB = new Date(b.year, b.month - 1, b.day);
    return dateA - dateB;
  });
  
  const firstDate = sortedDates[0];
  const lastDate = sortedDates[sortedDates.length - 1];
  
  // Google Calendarの日付形式: YYYYMMDD
  const startDate = `${firstDate.year}${String(firstDate.month).padStart(2, '0')}${String(firstDate.day).padStart(2, '0')}`;
  const endDate = `${lastDate.year}${String(lastDate.month).padStart(2, '0')}${String(lastDate.day).padStart(2, '0')}`;
  
  // カレンダーURLを構築（日付範囲を指定）
  const calendarUrl = `https://calendar.google.com/calendar/embed?src=${encodeURIComponent(BOT_CONFIG.CALENDAR_ID)}&dates=${startDate}/${endDate}`;
  
  // 日付の数に応じてサブタイトルを変更
  let subtitle;
  if (dates.length === 1) {
    subtitle = `対象日: ${firstDate.year}/${firstDate.month}/${firstDate.day}`;
  } else {
    subtitle = `対象日: ${firstDate.year}/${firstDate.month}/${firstDate.day} - ${lastDate.year}/${lastDate.month}/${lastDate.day}`;
  }
  
  return {
    'header': {
      'title': '📅 CI部-勤怠管理カレンダー',
      'subtitle': subtitle,
      'imageUrl': 'https://www.gstatic.com/images/branding/product/1x/calendar_48dp.png'
    },
    'sections': [
      {
        'widgets': [
          {
            'keyValue': {
              'topLabel': 'カレンダー',
              'content': '報告対象日のカレンダーを開く',
              'contentMultiline': true,
              'button': {
                'textButton': {
                  'text': 'カレンダーを開く',
                  'onClick': {
                    'openLink': {
                      'url': calendarUrl
                    }
                  }
                }
              }
            }
          }
        ]
      }
    ]
  };
}
