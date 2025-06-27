function dupulicateSpreadSheet() {
  // 設定情報を保持するスプレッドシートから情報を取得
  const CONFIG_SHEET_ID = '1nEg6VHiLmdhqUEzk5fFqQgXdgABSaznr9jLWO70pQ6g';  // info シートがあるスプレッドシートID
  const configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  
  // システム設定値の取得
  const configSheet = configSS.getSheetByName('config');
  const configData = configSheet.getDataRange().getValues();
  
  // 設定値をオブジェクトとして保持
  const config = {
    targetFolderId: '',
    templateFileId: '',
    defaultCiDeptEmail: ''
  };
  
  // configシートから設定値を読み込み
  for (let i = 1; i < configData.length; i++) {
    const key = configData[i][0];
    const value = configData[i][1];
    if (key in config) {
      config[key] = value;
    }
  }

  // 各種オブジェクトの取得
  const targetFolder = DriveApp.getFolderById(config.targetFolderId);
  const sourceFile = DriveApp.getFileById(config.templateFileId);
  
  // 個人情報シートの取得
  const infoSheet = configSS.getSheetByName('info');
  const data = infoSheet.getDataRange().getValues();

  // 個人情報シートから設定パラメータ取得
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '') {
      let fullName = "氏名：　" + data[i][0],
          fileName = data[i][1],
          mailaddr = data[i][2];
      
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
        targetSS.getRange("B4").setValue(fullName);
        targetSSFile.addEditor(mailaddr);
        
        // CI部門のメールアドレスも編集者として追加
        if (config.defaultCiDeptEmail) {
          targetSSFile.addEditor(config.defaultCiDeptEmail);
        }
      }
    }
  }
}
    