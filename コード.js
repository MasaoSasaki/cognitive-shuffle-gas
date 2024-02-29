let createdFilesCount = 0;
let existingFilesCount = 0;

// ルートフォルダの取得
const folderId = "1FjQAlZKzoDafr0g4-tog48HkVHlJvYwo";
const rootFolder = DriveApp.getFolderById(folderId);

function requestTextToSpeechForSpreadsheetWords() {
  const spreadsheetId = '1zJXW5c0G25oJ9CzlZ1Zx1IA1iiHODKVSstyB_alvnOg'; // ここにスプレッドシートのIDを設定
  const sheet = SpreadsheetApp.openById(spreadsheetId);

  // 実際にデータが入力されている最終行を取得
  const lastRow = sheet.getLastRow();
  const words = sheet.getRange(`B2:B${lastRow}`).getValues().map(row => row[0]);
  const names = sheet.getRange('D1:E1').getValues()[0];

  // ルートフォルダ内を空にする
  clearFolder();

  for (const name of names) {
    for (const word of words) {
      if (word) {
        textToSpeech(word, name);
      } else {
        break;
      }
    };
  }

  const totalCount = createdFilesCount + existingFilesCount;
  console.info(`作成されたファイル数: ${createdFilesCount}`);
  console.info(`既に存在していたファイル数: ${existingFilesCount}`);
  console.info(`合計処理ファイル数: ${totalCount}`);
}

function textToSpeech(word, name) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const gcpApiKey = scriptProperties.getProperty('GCP_API_KEY');
  const apiUrl = "https://texttospeech.googleapis.com/v1beta1/text:synthesize?key=" + gcpApiKey;

  // 保存先のフォルダとファイル名の設定
  const folderName = name; // 性別に基づいたフォルダ名
  const fileName = `${word}.mp3`;
  
  // 性別に基づくサブフォルダの取得または作成
  let genderFolder;
  const folders = rootFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    genderFolder = folders.next();
  } else {
    genderFolder = rootFolder.createFolder(folderName);
  }

  // 性別フォルダ内で同名のファイルを検索
  const files = genderFolder.getFilesByName(fileName);
  if (files.hasNext()) {
    existingFilesCount += 1;
    console.warn(`${folderName}/${fileName}は既に存在しています。`);
    return;
  } else {
    console.info(`${folderName}/${fileName}の新規作成を開始します。`);
  }

  const payload = JSON.stringify({
    "input": {
      "text": word,
    },
    "voice": {
      "languageCode": "ja-JP",
      "name": name,
    },
    "audioConfig": {
      "audioEncoding": "MP3",
    }
  });

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.audioContent) {
      createdFilesCount += 1;
      const decodedAudioContent = Utilities.base64Decode(jsonResponse.audioContent);
      const blob = Utilities.newBlob(decodedAudioContent, 'audio/mpeg', `${word}.mp3`);
      genderFolder.createFile(blob);
      console.info("音声ファイルがDriveに保存されました。");
    } else {
      console.error("エラーが発生しました。レスポンス: " + response.getContentText());
    }
  } catch (e) {
    console.error("例外が発生しました: " + e.toString());
  }
}

function clearFolder() {
  const files = rootFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.setTrashed(true);
  }
  
  const subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    subFolder.setTrashed(true);
  }
}
