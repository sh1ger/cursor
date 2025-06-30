# CI-休暇管理ボット 仕様書・利用手順書

## 概要

CI-休暇管理ボットは、Google Calendarの休暇申請イベントを自動監視し、申請・取消の変更を検知して人事担当者に自動通知するGoogle Apps Scriptシステムです。

### 主な機能
- Google Calendarの定期チェックによる休暇連絡集計
- 申請・取消の自動検知
- 対象者ごとの集計報告メール送信
- カスタムFromアドレスでのメール送信
- エラー発生時の自動通知

## システム構成

### 技術スタック
- **プラットフォーム**: Google Apps Script
- **連携サービス**: Google Calendar, Gmail
- **データ保存**: Script Properties（実行状態・イベント状態）

### ファイル構成
```
CI-休暇管理/
├── calendar-notification.js  # メインスクリプト
└── appsscript.json          # プロジェクト設定
```

## 詳細仕様

### 1. 監視対象カレンダー
- **カレンダーID**: `your-calendar-id@group.calendar.google.com`
- **タイムゾーン**: Asia/Tokyo
- **検索期間**: 過去1日〜未来30日

### 2. 対象となる休暇種別
- 全休
- 午前休
- 午後休
- 遅刻
- 早退
- 特別休
- 休出

### 3. イベント形式要件
カレンダーイベントは以下の形式である必要があります：

**タイトル形式**: `[対象者名] - [休暇種別]`
例: `田中太郎 - 全休`

**説明文形式**: 
```
申請者: [申請者名]
[その他の情報]
```

### 4. 通知設定
- **人事担当者メール**: hr@yourcompany.com
- **管理者メール**: admin@yourcompany.com
- **送信者アドレス**: attendance@yourcompany.com
- **送信者名**: CI-休暇管理通知
- **件名プレフィックス**: 【勤怠連絡】

### 5. 実行スケジュール
- **定期実行**: 毎日 9:00 と 17:00
- **手動実行**: テスト用関数あり

## 利用手順書

### 初期設定

#### 1. Google Apps Scriptプロジェクトの作成
1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を「CI-休暇管理」に設定

#### 2. スクリプトの配置
1. `calendar-notification.js` の内容をメインエディタにコピー
2. ファイルを保存（Ctrl+S）

#### 3. 権限の設定
初回実行時に以下の権限を許可：
- Google Calendar へのアクセス
- Gmail の送信権限
- Script Properties へのアクセス

#### 4. 設定値の確認・変更
`CONFIG` オブジェクト内の設定値を確認し、必要に応じて変更：

```javascript
const CONFIG = {
  CALENDAR: {
    ID: 'your-calendar-id@group.calendar.google.com',  // 監視対象カレンダーID
    TIMEZONE: 'Asia/Tokyo',
    SEARCH_DAYS_FORWARD: 30,
    DEFAULT_SEARCH_DAYS_BACK: 1
  },
  NOTIFICATION: {
    HR_EMAIL: 'hr@yourcompany.com',      // 人事担当者メール
    ADMIN_EMAIL: 'admin@yourcompany.com', // 管理者メール
    EMAIL_OPTIONS: {
      from: 'attendance@yourcompany.com',    // 送信者アドレス
      name: 'CI-休暇管理通知'         // 送信者名
    }
  }
};
```

### 運用開始

#### 1. トリガーの設定
```javascript
// スクリプトエディタで実行
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
```

#### 2. 初期データのクリア（初回実行時）
```javascript
// スクリプトエディタで実行
function clearStoredData() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('last_execution_time');
  properties.deleteProperty('previous_event_state');
  console.log('保存データをクリアしました');
}
```

#### 3. 設定確認
```javascript
// スクリプトエディタで実行
function checkConfiguration() {
  console.log('=== 設定確認 ===');
  console.log('カレンダーID:', CONFIG.CALENDAR.ID);
  console.log('人事メール:', CONFIG.NOTIFICATION.HR_EMAIL);
  // その他の設定値も表示
}
```

### テスト実行

#### 1. 手動テスト実行
```javascript
// スクリプトエディタで実行
function testAttendanceReport() {
  executeAttendanceReport();
}
```

#### 2. ログの確認
- 実行ログは「実行」→「実行ログ」で確認可能
- エラーが発生した場合は詳細が表示される

### 運用・保守

#### 1. 定期実行の監視
- トリガー設定は「トリガー」メニューで確認可能
- 実行履歴は「実行」→「実行履歴」で確認可能

#### 2. エラー対応
エラーが発生した場合：
1. 実行ログでエラー内容を確認
2. 設定値（カレンダーID、メールアドレス等）を確認
3. 必要に応じて権限を再設定

#### 3. データのリセット
問題が発生した場合のデータリセット：
```javascript
function clearStoredData() {
  // 上記のクリア関数を実行
}
```

## メール通知形式

### 通常通知メール
**件名**: 【勤怠連絡】CI部_YYYY/MM/DD-HH時

**本文例**:
```
CI部の勤怠連絡です。

【申請】
  田中太郎: 全休 (2025/1/15)
  佐藤花子: 午前休 (2025/1/16)

【取消】
  山田次郎: 全休 (2025/1/17)

---
このメールは「CI-休暇管理カレンダー」の自動通知メールです。

カレンダー: https://calendar.google.com/calendar/embed?src=...
```

### エラー通知メール
**件名**: 【勤怠連絡】 エラー通知 - [エラータイプ]

**本文例**:
```
休暇管理システムでエラーが発生しました。

【エラー詳細】
タイプ: ATTENDANCE_REPORT
エラー: カレンダーが見つかりません
発生時刻: 2025/1/15 9:00:00

---
このメールは休暇管理システムにより自動送信されています。
```

## トラブルシューティング

### よくある問題と対処法

#### 1. カレンダーが見つからないエラー
**原因**: カレンダーIDが間違っている、またはアクセス権限がない
**対処法**: 
- カレンダーIDを確認
- カレンダーの共有設定を確認
- スクリプトの権限を再設定

#### 2. メール送信エラー
**原因**: 送信者アドレスの設定が間違っている
**対処法**:
- 送信者アドレスが有効か確認
- Gmail送信権限を再設定

#### 3. イベントが検知されない
**原因**: イベントのタイトル形式が正しくない
**対処法**:
- イベントタイトルが「[名前] - [種別]」形式か確認
- 対象種別に含まれているか確認

#### 4. 重複通知が発生する
**原因**: 前回実行状態のデータが破損している
**対処法**:
```javascript
clearStoredData(); // データをクリアして再実行
```

## セキュリティ・プライバシー

### データ保護
- 個人情報はGoogle Apps Scriptの安全な環境で処理
- 外部へのデータ送信は通知メールのみ
- 実行ログはGoogleアカウント内でのみ保存

### アクセス制御
- スクリプトの編集権限は管理者のみ
- カレンダーアクセス権限は必要最小限
- メール送信権限は指定アドレスのみ

## 更新履歴

- **v1.0.0** (2025-06-27): 初回リリース
  - 基本的な休暇申請監視機能
  - 自動通知機能
  - エラー処理機能

## サポート・連絡先

システムに関する質問や問題がございましたら、以下までご連絡ください：
- **開発者**: システム管理者
- **連絡先**: admin@yourcompany.com

---

*このドキュメントは CI-休暇管理ボット v1.0.0 に基づいて作成されています。* 