let createdFilesCount = 0;
let existingFilesCount = 0;

function requestTextToSpeechForSpreadsheetWords() {
  const spreadsheetId = '1zJXW5c0G25oJ9CzlZ1Zx1IA1iiHODKVSstyB_alvnOg'; // ここにスプレッドシートのIDを設定
  const sheet = SpreadsheetApp.openById(spreadsheetId);

  // 実際にデータが入力されている最終行を取得
  const lastRow = sheet.getLastRow();
  const range = `B2:B${lastRow}`;
  const values = sheet.getRange(range).getValues();

  values.forEach(function(row) {
    const word = row[0];
    if (word) { // ワードが空でない場合のみリクエストを行う
      textToSpeech(word);
    }
  });

  const totalCount = createdFilesCount + existingFilesCount;
  console.info(`作成されたファイル数: ${createdFilesCount}`);
  console.info(`既に存在していたファイル数: ${existingFilesCount}`);
  console.info(`合計処理ファイル数: ${totalCount}`);
}

function textToSpeech(word) {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Firebaseプロジェクトの設定
  const firebaseApiKey = scriptProperties.getProperty('FIREBASE_API_KEY');
  const storageBucketUrl = 'cognitive-shuffle-415210.appspot.com';

  const fileName = `${word}.mp3`;
  const filePath = `voices/woman/${fileName}`;

  const gcpApiKey = scriptProperties.getProperty('GCP_API_KEY');
  const apiUrl = "https://texttospeech.googleapis.com/v1/text:synthesize?key=" + gcpApiKey;

  // const fileName = `${word}.mp3`;
  // const folderId = "1FjQAlZKzoDafr0g4-tog48HkVHlJvYwo"; // ここにフォルダのIDを設定

  // // Google Driveで同名のファイルを検索
  // const folder = DriveApp.getFolderById(folderId);
  // const files = folder.getFilesByName(fileName);
  // if (files.hasNext()) {
  //   existingFilesCount += 1;
  //   console.warn(`${fileName}は既に存在しています。`);
  //   return;
  // } else {
  //   console.info(`${fileName}の新規作成を開始します。`);
  // }

  const payload = JSON.stringify({
    "input": {
      "text": word,
    },
    "voice": {
      "languageCode": "ja-JP",
      "name": "ja-JP-Wavenet-A",
      "ssmlGender": "FEMALE"
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
      // Firebase Storageへのアップロード
      const uploadUrl = `https://firebasestorage.googleapis.com/v0/b/${storageBucketUrl}/o?name=${encodeURIComponent(filePath)}`;
      const uploadOptions = {
        "method": "put",
        "contentType": "application/octet-stream",
        "headers": {
          "Authorization": "Bearer " + firebaseApiKey,
        },
        "payload": Utilities.base64Decode(jsonResponse.audioContent)
      };

      const uploadResponse = UrlFetchApp.fetch(uploadUrl, uploadOptions);
      const uploadResponseJson = JSON.parse(uploadResponse.getContentText());

      if (uploadResponse.getResponseCode() === 200) {
        console.info("音声ファイルがFirebase Storageに保存されました。");
      } else {
        console.error("Firebase Storageへの保存に失敗しました。レスポンス: " + uploadResponse.getContentText());
      }
    } else {
      console.error("エラーが発生しました。レスポンス: " + response.getContentText());
    }
    // if (jsonResponse.audioContent) {
    //   createdFilesCount += 1;
    //   const decodedAudioContent = Utilities.base64Decode(jsonResponse.audioContent);
    //   const blob = Utilities.newBlob(decodedAudioContent, 'audio/mpeg', `${word}.mp3`);
    //   folder.createFile(blob);
    //   console.info("音声ファイルがDriveに保存されました。");
    // } else {
    //   console.error("エラーが発生しました。レスポンス: " + response.getContentText());
    // }
  } catch (e) {
    console.error("例外が発生しました: " + e.toString());
  }
}
