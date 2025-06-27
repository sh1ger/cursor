function getConfig() {
  // 設定値とinfo情報を保持する共通のスプレッドシートを開く
  const MASTER_SHEET_ID = '1GLXSb8yDYTF_Vtd_R_j-ligo6248WZy1dB4lMNajljo'; // このIDは初回のみハードコード
  const masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const configSheet = masterSS.getSheetByName('config');
  
  if (!configSheet) {
    throw new Error('設定シート「config」が見つかりません');
  }

  const configData = configSheet.getDataRange().getValues();
  const config = {};
  
  // 設定値をオブジェクトに格納
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0]) {
      config[configData[i][0]] = configData[i][1];
    }
  }
  
  // 必須の設定値をチェック
  const requiredKeys = ['targetFolderId', 'templateFileId', 'ciDeptEmail'];
  for (const key of requiredKeys) {
    if (!config[key]) {
      throw new Error(`必須の設定値「${key}」が設定されていません`);
    }
  }
  
  return {
    config: config,
    spreadsheet: masterSS
  };
}

function dupulicateSpreadSheet() {
  try {
    // 設定値とスプレッドシートオブジェクトを取得
    const { config, spreadsheet } = getConfig();

    // 目標設定シート作成先フォルダ（半期毎）
    const targetFolder = DriveApp.getFolderById(config.targetFolderId);

    // 情報シートの情報
    const infoSheet = spreadsheet.getSheetByName('info');
    if (!infoSheet) {
      throw new Error('情報シート「info」が見つかりません');
    }
    const data = infoSheet.getDataRange().getValues();

    // テンプレートファイルの情報
    const fileId = config.templateFileId,
          sourceFile = DriveApp.getFileById(fileId);

    // 情報シートから設定パラメータ取得
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== '') {
        let fullName = data[i][0],
            fileName = data[i][1],
            age = data[i][3],
            exp = data[i][4],
            mailaddr = data[i][2],
            cidept = config.ciDeptEmail;
        
        // 目標設定シートファイルのコピー
        sourceFile.makeCopy(fileName, targetFolder);

        // コピーしたファイルオブジェクトの取得
        let targetFileObj = targetFolder.getFilesByName(fileName);
        while (targetFileObj.hasNext()) {
          let targetFile = targetFileObj.next(),
              targetFileId = targetFile.getId(),
              targetSSFile = SpreadsheetApp.openById(targetFileId),
              targetSS = targetSSFile.getSheets()[0];

          // ファイル内セルに個人情報をセット
          targetSS.getRange("E4").setValue(fullName);
          targetSS.getRange("Q3").setValue(age);
          targetSS.getRange("Q4").setValue(exp);

          // 上記でセットした値を保護
          let protectRange = targetSS.getRange("C3:V4"),
              protections = protectRange.protect(),
              userList = protections.getEditors();
              
          // オーナー以外の編集権限を剥奪    
          protections.removeEditors(userList);

          // ファイル共有権限の追加
          targetSSFile.addEditor(mailaddr);
          targetSSFile.addViewer(cidept);
        }
      }
    }
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message);
    throw error;
  }
}
