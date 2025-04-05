// Twitter API関連URL
const ACCESS_TOKEN_URL = "https://api.twitter.com/2/oauth2/token";
const REQUEST_TOKEN_URL = "https://api.twitter.com/2/oauth2/authorize";
const AUTHORIZE_URL = "https://twitter.com/i/oauth2/authorize";
const X_POST_URL_TEMPLATE = "https://x.com/nishi_hiko117/status/1865580350833119380"
const SERVICE_NAME = "twitter";

/**
 * アカウント認証用リンクを表示
 */
function authorizeLinkForNewTwitterBotWithImage() {
  const service = getXService();

  // システムプロパティを取得
  const prop = getSystemProperty();


  if (!service) {
    return;
  }
  const ui = SpreadsheetApp.getUi();

  if (!service.hasAccess()) {
    const authorizationURL = getAuthorizationUrl();
    const template = HtmlService.createTemplateFromFile("XAuthorization");
    template.authorizationURL = authorizationURL;
    const html = template.evaluate().setWidth(600).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, "Twitterアカウント認証");
  } else {
    ui.alert("アカウント認証はすでに許可されています。");
  }
}

/**
 * 認証コールバック関数
 * @param {Object} request eパラメータ
 * @returns {HtmlOutput} 認証結果HTML出力
 */
function authCallback(request) {
  
  const prop = getSystemProperty();

  // リクエストパラメータのログ出力
  Logger.log('コールバックリクエスト: ' + JSON.stringify(request.parameter));
  
  var service = getXService();
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var codeVerifier = getSystemPropertyValue(PROPERTY_CELL.X_CODE_VERIFIER);
  var clientId = prop.xApiClient;
  var clientSecret = prop.xApiClientSecret;
  var redirectUri = getRedirectUri();
  
  Logger.log('コールバック処理 - code_verifier: ' + codeVerifier);
  Logger.log('リダイレクトURI: ' + redirectUri);
  
  // code_verifierを直接URLパラメーターとして追加
  var payload = {
    'code': request.parameter.code,
    'code_verifier': codeVerifier,
    'grant_type': 'authorization_code',
    'redirect_uri': redirectUri,
    'client_id': clientId
  };
  
  // トークンリクエストのオプションを設定
  var tokenOptions = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payload,
    'headers': {
      'Authorization': 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret)
    },
    'muteHttpExceptions': true
  };
  
  try {
    // リクエスト情報をログ出力（デバッグ用）
    Logger.log('トークンリクエスト URL: https://api.twitter.com/2/oauth2/token');
    Logger.log('トークンリクエスト ペイロード: ' + JSON.stringify(payload));
    Logger.log('トークンリクエスト ヘッダー: ' + JSON.stringify(tokenOptions.headers));
    
    // 直接トークンをリクエスト
    var response = UrlFetchApp.fetch('https://api.twitter.com/2/oauth2/token', tokenOptions);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    Logger.log('トークンレスポンス コード: ' + responseCode);
    Logger.log('トークンレスポンス: ' + responseText);
    
    // レスポンスのパース
    if (responseCode >= 200 && responseCode < 300) {
      try {
        var tokenData = JSON.parse(responseText);
        
        if (tokenData.access_token) {
          // トークンを保存
          setSystemPropertyValue(PROPERTY_CELL.X_OAUTH2_TWITTER, JSON.stringify({
            access_token: tokenData.access_token,
            refresh_token: tokenData.refresh_token,
            expires_in: tokenData.expires_in,
            timestamp: new Date().getTime()
          }));
          
          return HtmlService.createHtmlOutput(
            '<h3>認証が成功しました</h3>' +
            '<p>このタブを閉じて、スクリプトに戻ってください。</p>' +
            '<p>アクセストークンが正常に取得されました。</p>'
          );
        } else {
          return HtmlService.createHtmlOutput(
            '<h3>認証エラー</h3>' +
            '<p>アクセストークンが取得できませんでした。</p>' +
            '<pre>' + JSON.stringify(tokenData, null, 2) + '</pre>'
          );
        }
      } catch (parseError) {
        Logger.log('JSONパースエラー: ' + parseError);
        return HtmlService.createHtmlOutput(
          '<h3>レスポンス解析エラー</h3>' +
          '<p>APIからのレスポンスを解析できませんでした。</p>' +
          '<p>エラー: ' + parseError + '</p>' +
          '<p>レスポンス: ' + responseText + '</p>'
        );
      }
    } else {
      // エラーレスポンスの処理
      return HtmlService.createHtmlOutput(
        '<h3>APIエラー</h3>' +
        '<p>ステータスコード: ' + responseCode + '</p>' +
        '<p>エラーメッセージ: ' + responseText + '</p>' +
        '<p>このエラーについては、X Developer Portalの設定を確認してください。</p>' +
        '<p>特に、リダイレクトURI: <code>' + getRedirectUri() + '</code> が正しく設定されていることを確認してください。</p>'
      );
    }
  } catch (e) {
    Logger.log('トークン取得エラー: ' + e.toString());
    return HtmlService.createHtmlOutput(
      '<h3>リクエストエラー</h3>' +
      '<p>APIへのリクエスト中にエラーが発生しました。</p>' +
      '<p>エラー: ' + e.toString() + '</p>' +
      '<p>以下の点を確認してください：</p>' +
      '<ul>' +
      '<li>スクリプトプロパティにCLIENT_IDとCLIENT_SECRETが正しく設定されているか</li>' +
      '<li>X Developer PortalでリダイレクトURLが <code>' + getRedirectUri() + '</code> として登録されているか</li>' +
      '<li>TwitterアプリでOAuth 2.0とPKCEが有効になっているか</li>' +
      '</ul>'
    );
  }
}

/**
 * 認証情報からユーザーIDをセットする
 */
function setUserId(){
  let userId = getUserIdFromApiKey();

  // システムプロパティを記録する
  setSystemPropertyValue(PROPERTY_CELL.X_USER_ID, userId);  // XユーザーID
  
  const ui = SpreadsheetApp.getUi();
  ui.alert("ユーザーIDの設定が完了しました。");
  Logger.log('ユーザーIDの設定が完了しました。');
}

//********************************
//　ポスト情報送信用処理
//********************************
function getXService() {
  // システムプロパティを取得
  const prop = getSystemProperty();
  
  // スプレッドシートIDを取得
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // アクセストークンがすでに保存されているか確認
  var tokenData = getSystemPropertyValue(PROPERTY_CELL.X_OAUTH2_TWITTER);
  Logger.log('保存されているトークンデータ: ' + tokenData);

  if (tokenData) {
    try {
      tokenData = JSON.parse(tokenData);
      if (tokenData.access_token) {
        var service = {
          hasAccess: function() { return true; },
          getAccessToken: function() { return tokenData.access_token; },
          reset: function() { 
            setSystemPropertyValue(PROPERTY_CELL.X_OAUTH2_TWITTER, "");
            setSystemPropertyValue(PROPERTY_CELL.X_CODE_VERIFIER, "");
          }
        };
        return service;
      }
    } catch (e) {
      Logger.log('トークンデータの解析エラー: ' + e);
    }
  }
  
  // 通常のサービス作成（認証前または再認証が必要な場合）
  return OAuth2.createService('twitter.' + spreadsheetId)  // サービス名にもスプレッドシートIDを含める
    .setAuthorizationBaseUrl('https://twitter.com/i/oauth2/authorize')
    .setTokenUrl('https://api.twitter.com/2/oauth2/token')
    .setClientId(prop.xApiClient)
    .setClientSecret(prop.xApiClientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())  // UserPropertiesからScriptPropertiesに変更
    .setScope('tweet.read tweet.write users.read offline.access')
    .setParam('response_type', 'code')
    .setParam('code_challenge_method', 'S256')
    .setRedirectUri(getRedirectUri())
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(
        prop.xApiClient + ':' + 
        prop.xApiClientSecret
      ),
      'Content-Type': 'application/x-www-form-urlencoded'
    });
}


// PKCE用のcode_verifierを生成
function generateCodeVerifier() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';
  var verifier = '';
  for (var i = 0; i < 128; i++) {
    verifier += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return verifier;
}

// PKCE用のcode_challengeを生成
function generateCodeChallenge(verifier) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, verifier);
  var encoded = Utilities.base64Encode(rawHash);
  // Base64 URL safe対応
  return encoded.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// リダイレクトURIを取得する関数
function getRedirectUri() {
  // スクリプトIDの取得
  var scriptId = ScriptApp.getScriptId();
  // リダイレクトURIの組み立て
  return 'https://script.google.com/macros/d/' + scriptId + '/usercallback';
}

// 認証URLを取得するための関数
function getAuthorizationUrl() {
  // まず古い認証情報をクリア
  resetAuth();

  // システムプロパティを取得
  const prop = getSystemProperty();
  
  // サービスを取得
  var service = getXService();
  
  // PKCE用のcode_challengeを生成
  var codeVerifier = generateCodeVerifier();
  var codeChallenge = generateCodeChallenge(codeVerifier);
  
  // code_verifierを保存（デバッグ用にログも出力）
  setSystemPropertyValue(PROPERTY_CELL.X_CODE_VERIFIER, codeVerifier);
  Logger.log('生成されたcode_verifier: ' + codeVerifier);
  Logger.log('生成されたcode_challenge: ' + codeChallenge);
  
  // code_challengeを設定
  service.setParam('code_challenge', codeChallenge);
  
  var authUrl = service.getAuthorizationUrl();
  Logger.log('認証URLを開いてください: %s', authUrl);
  return authUrl;
}

// 認証状態をリセット（トラブルシューティング用）
function resetAuth() {
  // ユーザープロパティから認証情報を削除
  setSystemPropertyValue(PROPERTY_CELL.X_OAUTH2_TWITTER, "");
  setSystemPropertyValue(PROPERTY_CELL.X_CODE_VERIFIER, "");
  
  Logger.log('認証状態をリセットしました。');
  return '認証状態をリセットしました。';
}

// 複数の画像をアップロードする関数
function uploadImagesForX(attachmentInfo, index) {
  
  const uploadMediaEndpoint = 'https://upload.twitter.com/1.1/media/upload.json';
  
  // 認証情報を取得する
  const service = getXService();

  Logger.log('X画像アップロード中:' + index + '件目...');

  Logger.log('=== 画像アップロード開始 ===');
  Logger.log(`インデックス: ${index}`);
  Logger.log(`URL: ${attachmentInfo.url}`);
  Logger.log(`FileID: ${attachmentInfo.fileId}`);

  try {
    // const imageBlob = UrlFetchApp.fetch(attachmentInfo.url).getBlob();

    let imageBlob;
    
    if (attachmentInfo.fileId) {
      const file = DriveApp.getFileById(attachmentInfo.fileId);
      imageBlob = file.getBlob();
      Logger.log(`Drive File Name: ${file.getName()}`);
      Logger.log(`Drive File MimeType: ${file.getMimeType()}`);
    } else {
      const response = UrlFetchApp.fetch(attachmentInfo.url);
      Logger.log('Response Headers:');
      Logger.log(JSON.stringify(response.getAllHeaders()));
      imageBlob = response.getBlob();
    }

    // 画像の詳細情報を出力
    Logger.log(`Blob ContentType: ${imageBlob.getContentType()}`);
    Logger.log(`Blob Name: ${imageBlob.getName()}`);
    Logger.log(`Blob Size: ${imageBlob.getBytes().length} bytes`);

    // MIMEタイプの検証
    const validMimeTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    const currentMimeType = imageBlob.getContentType();


    if (!validMimeTypes.includes(currentMimeType)) {
      Logger.log(`警告: 不適切なMIMEタイプ: ${currentMimeType}`);
      // MIMEタイプを強制的に設定
      imageBlob.setContentType('image/jpeg');
      Logger.log(`MIMEタイプを image/jpeg に変更しました`);
    }

    const imageBytes = imageBlob.getBytes();
    Logger.log(`Image Bytes Length: ${imageBytes.length}`);

    // JPEGシグネチャーのチェック
    const isJpeg = imageBytes[0] === 0xFF && imageBytes[1] === 0xD8;
    Logger.log(`JPEG Signature Check: ${isJpeg}`);

    // Base64エンコード前にバイナリデータを確認
    const base64Data = Utilities.base64Encode(imageBytes);
    Logger.log(`Base64 Length: ${base64Data.length}`);

    const payload = {
      media_data: base64Data,
      media_category: 'tweet_image'  // カテゴリを明示的に指定
    };

    
    const uploadResponse = service.fetch(uploadMediaEndpoint, {
      method: "POST",
      // payload: { 
      //   media_data: base64Data 
      // },
      payload: payload,
      contentType: "application/x-www-form-urlencoded",
      muteHttpExceptions: true
    });

    // X APIの制限に対応するため、アップロード間に少し待機
    Utilities.sleep(2000);

    // レスポンスの詳細なログ出力
    Logger.log('Response Code: ' + uploadResponse.getResponseCode());
    Logger.log('Response Content: ' + uploadResponse.getContentText());

    if (uploadResponse.getResponseCode() !== 200) {
      Logger.log(`X画像アップロード失敗: ${attachmentInfo.url}. Response: ${uploadResponse.getContentText()}`);
      throw new Error(`X画像アップロード失敗: ${attachmentInfo.url}. Response: ${uploadResponse.getContentText()}`);
    }

    Logger.log('X画像アップロード成功');

    // 直接media_id_stringを返す
    const mediaData = JSON.parse(uploadResponse);
    return mediaData.media_id_string;

  } catch (error) {
    Logger.log(`X画像アップロード失敗: ${attachmentInfo.url}: ${error.toString()}`);
    throw error;
  }
  
}

/**
 * 画像のアップロード
 * @returns {Array} media_idの配列
 */
function uploadImages() {
  const service = getXService();
  if (service === null) {
    return null;
  }

  const uploadedMediaIds = [];
  const uploadUrl = "https://upload.twitter.com/1.1/media/upload.json";

  for (let i = 0; i < this.tweetInfo.imageURLs.length; i++) {
    const imageUrl = this.tweetInfo.imageURLs[i];
    if (imageUrl === "") {
      continue;
    }

    const fileId = imageUrl.split("/d/")[1].split("/")[0];
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());

    const response = service.fetch(uploadUrl, {
      method: "POST",
      payload: { media_data: base64Data },
      muteHttpExceptions: true
    });

    console.log("uploadURLResponseCode:", response.getResponseCode());
    console.log("uploadURLResponse:", response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error(`画像のアップロードに失敗しました。画像URL：${imageUrl}：詳細:${response.getResponseCode()}:${response.getContentText()}`);
    }

    const json = JSON.parse(response);
    const mediaId = json.media_id_string;
    uploadedMediaIds.push(mediaId);
  }
  return uploadedMediaIds;
}

/**
 * 動画アップロード用の関数
 * @param {string} videoFile アップロードする動画情報
 * @returns {string} アップロードされた動画のmedia_id
 */
function uploadVideoForX(attachVideoInfo, index) {
  const service = getXService();
  if (!service) {
    throw new Error("サービスの取得に失敗しました。");
  }

  Logger.log('X動画アップロード中:' + index + '件目...');

  try {
    // Google Driveから動画ファイルを取得
    let videoFile;
    let videoBlob;
    let totalBytes;
    let videoInfo;
    let mimeType;

    if(attachVideoInfo.fileId != ""){
      // Google Driveから動画ファイルを取得
      videoFile = DriveApp.getFileById(attachVideoInfo.fileId);
      videoBlob = videoFile.getBlob();
      totalBytes = videoBlob.getBytes().length;
      mimeType = videoFile.getMimeType();

    }else{
      videoInfo = getVideoInfo(attachVideoInfo.url);
      videoBlob = videoInfo.blob;
      totalBytes = videoInfo.sizeBytes;
      mimeType = videoInfo.contentType;
    }

    // ファイルサイズチェック
    if (totalBytes > MAX_VIDEO_SIZE) {
      throw new Error(`動画ファイルサイズが制限（30MB）を超えています。現在のサイズ: ${Math.round(totalBytes / 1024 / 1024)}MB`);
    }

    // 動画ファイル形式チェック
    if (mimeType !== 'video/mp4') {
      throw new Error(`非対応の動画形式です。MP4形式のみ対応しています。現在の形式: ${mimeType}`);
    }

    Logger.log(`動画アップロード開始 - ファイルサイズ: ${Math.round(totalBytes / 1024 / 1024)}MB`);

    // STEP 1: 初期化
    const mediaId = initializeVideoUpload(service, totalBytes);
    
    // STEP 2: チャンク分割アップロード
    appendVideoChunks(service, mediaId, videoBlob);
    
    // STEP 3: アップロード完了通知
    finalizeVideoUpload(service, mediaId);
    
    // STEP 4: 処理完了待機
    waitForVideoProcessing(service, mediaId);
    
    Logger.log('動画アップロード完了');
    return mediaId;
  } catch (error) {
    Logger.log('動画アップロードエラー: ' + error.toString());
    throw error;
  }
}

/**
 * 動画アップロードの初期化
 */
function initializeVideoUpload(service, totalBytes) {
  const endpoint = 'https://upload.twitter.com/1.1/media/upload.json';
  const payload = {
    command: 'INIT',
    total_bytes: totalBytes,
    media_type: 'video/mp4',
    media_category: 'tweet_video'
  };
  
  const response = service.fetch(endpoint, {
    method: 'POST',
    payload: payload,
    muteHttpExceptions: true
  });
  
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  const responseData = JSON.parse(responseText);
  
  Logger.log('初期化レスポンス:', responseText);
  
  // 202も正常なレスポンスとして扱う
  if (responseCode !== 200 && responseCode !== 202) {
    throw new Error('動画アップロードの初期化に失敗しました: ' + responseText);
  }
  
  if (!responseData.media_id_string) {
    throw new Error('media_id_stringが取得できませんでした: ' + responseText);
  }
  
  Logger.log('動画アップロードの初期化成功 - media_id: ' + responseData.media_id_string);
  return responseData.media_id_string;
}

/**
 * 動画データのチャンク分割アップロード（安定化版）
 */
function appendVideoChunks(service, mediaId, videoBlob) {
  const chunkSize = 5 * 1024 * 1024; // 5MB chunks
  const totalBytes = videoBlob.getBytes().length;
  const chunks = Math.ceil(totalBytes / chunkSize);
  
  Logger.log('========== チャンクアップロード開始 ==========');
  Logger.log(`総バイト数: ${totalBytes} bytes (${Math.round(totalBytes / 1024 / 1024)}MB)`);
  Logger.log(`チャンクサイズ: ${chunkSize} bytes (${Math.round(chunkSize / 1024 / 1024)}MB)`);
  Logger.log(`総チャンク数: ${chunks}`);

  // メモリ使用量を最適化するため、チャンク処理を分割
  const bytes = videoBlob.getBytes();
  for (let i = 0; i < chunks; i++) {
    const start = i * chunkSize;
    const end = Math.min(start + chunkSize, totalBytes);
    
    Logger.log(`\n----- チャンク ${i + 1}/${chunks} アップロード開始 -----`);
    Logger.log(`チャンクサイズ: ${end - start} bytes`);

    try {
      // チャンクを個別に処理
      const chunk = bytes.slice(start, end);
      const base64Data = Utilities.base64Encode(chunk);

      const payload = {
        command: 'APPEND',
        media_id: mediaId,
        segment_index: i,
        media_data: base64Data
      };

      Logger.log(`リクエスト準備完了 (media_id: ${mediaId}, segment_index: ${i})`);

      const response = service.fetch('https://upload.twitter.com/1.1/media/upload.json', {
        method: 'POST',
        payload: payload,
        muteHttpExceptions: true
      });

      const responseCode = response.getResponseCode();
      
      // 204（No Content）も成功として扱う
      if (responseCode !== 200 && responseCode !== 202 && responseCode !== 204) {
        const responseText = response.getContentText();
        const headers = response.getAllHeaders();
        throw new Error(
          `チャンクアップロード失敗:\n` +
          `ステータスコード: ${responseCode}\n` +
          `レスポンス: ${responseText}\n` +
          `ヘッダー: ${JSON.stringify(headers)}`
        );
      }

      Logger.log(`チャンク ${i + 1}/${chunks} アップロード成功 (ステータスコード: ${responseCode})`);

      // メモリを解放
      delete chunk;
      delete base64Data;

      // スクリプトの実行時間制限に対応するため、
      // 定期的にチェックポイントを設ける
      if (i > 0 && i % 5 === 0) {
        Logger.log('チェックポイント - 短い休止を入れます');
        Utilities.sleep(2000);
      }

    } catch (error) {
      Logger.log('\n===== エラー詳細 =====');
      Logger.log(error.toString());
      throw new Error(`チャンク ${i + 1}/${chunks} のアップロードに失敗しました: ${error.toString()}`);
    }

    // API制限を考慮して待機
    Logger.log('API制限対応のため待機');
    Utilities.sleep(1000);
  }

  // メモリを解放
  delete bytes;

  Logger.log('\n========== 全チャンクのアップロード完了 ==========');
}

/**
 * レスポンスの詳細なログを出力する補助関数
 */
function logDetailedResponse(response, context = '') {
  Logger.log(`\n===== レスポンス詳細 ${context} =====`);
  Logger.log(`ステータスコード: ${response.getResponseCode()}`);
  
  try {
    const contentText = response.getContentText();
    Logger.log('レスポンス本文:');
    Logger.log(contentText);
    
    try {
      const jsonResponse = JSON.parse(contentText);
      Logger.log('JSONパース結果:');
      Logger.log(JSON.stringify(jsonResponse, null, 2));
    } catch (e) {
      Logger.log('レスポンスのJSONパースに失敗:');
      Logger.log(e.toString());
    }
  } catch (e) {
    Logger.log('レスポンス本文の取得に失敗:');
    Logger.log(e.toString());
  }

  Logger.log('レスポンスヘッダー:');
  Logger.log(JSON.stringify(response.getAllHeaders(), null, 2));
  Logger.log('=====================================\n');
}

/**
 * 動画アップロードの完了通知
 */
function finalizeVideoUpload(service, mediaId) {
  const response = service.fetch('https://upload.twitter.com/1.1/media/upload.json', {
    method: 'POST',
    payload: {
      command: 'FINALIZE',
      media_id: mediaId
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    throw new Error('動画アップロードの完了処理に失敗しました: ' + response.getContentText());
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * 動画処理の完了待機
 */
function waitForVideoProcessing(service, mediaId) {
  const maxAttempts = 30; // 最大待機回数
  let attempts = 0;
  
  while (attempts < maxAttempts) {
    const response = service.fetch('https://upload.twitter.com/1.1/media/upload.json?command=STATUS&media_id=' + mediaId, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('ステータス確認に失敗しました: ' + response.getContentText());
    }
    
    const status = JSON.parse(response.getContentText());
    
    if (status.processing_info) {
      if (status.processing_info.state === 'succeeded') {
        return true;
      } else if (status.processing_info.state === 'failed') {
        throw new Error('動画の処理に失敗しました: ' + JSON.stringify(status.processing_info.error));
      }
    }
    
    Logger.log(`動画処理待機中... 試行回数: ${attempts + 1}/${maxAttempts}`);
    attempts++;
    Utilities.sleep(2000); // 2秒待機
  }
  
  throw new Error('動画の処理がタイムアウトしました');
}

// 複数画像付きツイートを投稿する関数
function postTweetWithMultipleImages(tweetText, attachmentInfos, resId, quoteId) {
  try {
    // 認証情報を取得する
    const service = getXService();

    // 画像数の制限チェック
    if (attachmentInfos.length > 4) {
      throw new Error('Xへの投稿は最大4画像までとなります。');
    }

    // ツイートのペイロードを作成
    const tweetEndpoint = 'https://api.twitter.com/2/tweets';

    // 全画像をアップロード(ない場合は空の配列を返す)
    const mediaIds = [];
    for (const [index, attachmentInfo] of attachmentInfos.entries()) {
      let mediaId;
      if ( attachmentInfo.fileCategory == CONFIG.STRING_IMAGE){
        mediaId = uploadImagesForX(attachmentInfo, index + 1);
      } else {
        mediaId = uploadVideoForX(attachmentInfo, index + 1);
      }
      
     // 直接media_idを配列に追加
      mediaIds.push(mediaId);
    }

    // 返信・引用・通常ツイートを分類する
    if(resId != ""){
      // 返信IDを取得する

      // 返信ツイート
      payloadObj = {
        text: tweetText,
        reply: {
          in_reply_to_tweet_id: resId
        }
      };

    }else if(quoteId != ""){
      // 引用ツイート
      payloadObj = {
        text: tweetText,
        quote_tweet_id: quoteId
      };
    }else{
      // 通常ツイート
      payloadObj = {
        text: tweetText
      };
    }

    // // メディアIDをカンマ区切りでつなぐ
    // let mediaIdString = mediaIds.join(",");
    // // 添付ファイルがあれば添付する
    // if (mediaIds.length > 0) {
    //   payloadObj.media = { media_ids: mediaIdString };
    // }

    // 添付ファイルがあれば添付する
    if (mediaIds.length > 0) {
      payloadObj.media = { media_ids: mediaIds }; // 直接配列を使用
    }

    const options = {
      method: "post",
      payload: JSON.stringify(payloadObj),
      contentType: "application/json",
      muteHttpExceptions: true
    };

    // 投稿を実施
    const tweetId = makeTweetRequestForNewTwitterBotWithImage(tweetEndpoint, options);

    // ユーザーIDを取得
    let userId = getXUserId();

    return{
      tweetId: tweetId,
      url: getPostUrl(userId, tweetId)
    }

  } catch (error) {
    Logger.log('複数画像のX投稿失敗エラー内容: ' + error.toString());
    throw error;
  }
}

/**
 * 動画付きツイートを投稿する関数
 * @param {string} tweetText ツイート本文
 * @param {string} videoFile 動画ファイル情報
 * @returns {Object} 投稿結果（tweetIdとURL）
 */
function postTweetWithVideo(tweetText, videoFile) {
  try {
    // 認証情報を取得する
    const service = getXService();

    // 動画をアップロード
    const mediaId = uploadVideoForX(videoFile, 1);

    // ツイートのペイロードを作成
    const tweetEndpoint = 'https://api.twitter.com/2/tweets';
    const payloadObj = {
      text: tweetText,
      media: { media_ids: [mediaId] }
    };

    const options = {
      method: "post",
      payload: JSON.stringify(payloadObj),
      contentType: "application/json",
      muteHttpExceptions: true
    };

    // 投稿を実施
    const tweetId = makeTweetRequestForNewTwitterBotWithImage(tweetEndpoint, options);

    // ユーザーIDを取得
    let userId = getXUserId();

    return {
      tweetId: tweetId,
      url: getPostUrl(userId, tweetId)
    }

  } catch (error) {
    Logger.log('動画付きX投稿失敗エラー内容: ' + error.toString());
    throw error;
  }
}

// ユーザー情報の取得
function getUserIdFromApiKey() {
  const service = getXService();
  const url = "https://api.twitter.com/2/users/me"; // 認証されたユーザー情報を取得
  const options = {
    method: "get",
  };

  const response = service.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  
  if (!json.data) {
    throw new Error("ユーザー情報の取得に失敗しました: " + response.getContentText());
  }

  return json.data.username; // ユーザーIDを返す
}

/**
 * ツイート送信リクエスト用関数
 * @param {String} url エンドポイントURL
 * @param {Object} options fetchオプション
 * @returns {String} ツイートID
 */
function makeTweetRequestForNewTwitterBotWithImage(url, options) {
  const service = getXService();
  if (!service) {
    throw new Error("ツイートに失敗しました。");
  }

  const response = service.fetch(url, options);
  const responseCode = response.getResponseCode();
  const json = JSON.parse(response.getContentText());
  const jsonString = JSON.stringify(json);

  Logger.log('responseCode: ' + responseCode);
  Logger.log('jsonString: ' + jsonString);
  console.log();
  console.log();

  if (responseCode !== 201) {
    throw new Error(`ツイートに失敗しました。詳細：${jsonString}`);
  }
  return json.data.id;
}

// レスポンスの詳細をログ出力する補助関数
function logResponse(response) {
  Logger.log('Response Code: ' + response.getResponseCode());
  Logger.log('Response Headers: ' + JSON.stringify(response.getAllHeaders()));
  Logger.log('Response Content: ' + response.getContentText());
}

/**
 * サービスクリア（認証解除）
 */
function clearServiceForNewTwitterBotWithImage() {
  // アカウントステータスを更新する
  clearSnsAccountSettingStatus(CONFIG.CELL_SETTING_STATUS_X); // ステータス：設定済
  clearSnsCheck(CONFIG.CELL_SETTING_CHECKBOX_X);  // チェックボックスオン
  
  // 認証情報をクリア
  setSystemPropertyValue(PROPERTY_CELL.X_OAUTH2_TWITTER, "");
  setSystemPropertyValue(PROPERTY_CELL.X_CODE_VERIFIER, "");
}

/**
 * スクリプトID取得
 */
function getScriptIDForNewTwitterBotWithImage() {
  const scriptId = ScriptApp.getScriptId();
  
  // スプレッドシートを取得する
  ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // スクリプトIDを設定用シートへ反映する
  ss.getSheetByName(SHEETS_NAME.SETTING).getRange(CONFIG_SETTING.CELL_SCRIPT_ID).setValue(scriptId);
  
}

/**
 * XユーザーID取得
 */
function getXUserId(){
  // プロパティを取得
  const prop = getSystemProperty();
  return prop.xUserId;
}

/**
 * 投稿URLの生成
 * @param {string} userId XのユーザーID
 * @param {string} postId 投稿ID
 * @returns {string} 投稿URL
 */
function getPostUrl(userId, postId){
    // テンプレート形式のURLを作成
    const templateUrl = `https://x.com/${userId}/status/${postId}`;
    return templateUrl;
}


/**
 * TwitterのURLからツイートIDを抽出する関数
 * @param {string} url - TwitterのURL
 * @return {string|null} ツイートID または null（無効なURLの場合）
 */
function extractTweetId(url) {
  try {
    // URLが空の場合はnullを返す
    if (!url) return null;
    
    // URLからstatusの後の数字を抽出する正規表現
    const regex = /(?:twitter\.com|x\.com)\/\w+\/status\/(\d+)/;
    
    // URLからツイートIDを抽出
    const match = url.match(regex);
    
    // マッチした場合はツイートIDを、しない場合は空文字を返す
    return match ? match[1] : "";
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    return null;
  }
}

function getSystemProperty() {
  // スプレッドシートを取得する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // システムシートを取得する
  const systemSheet = ss.getSheetByName(SHEETS_NAME.SYSTEM);
  
  // プロパティを取得する
  const prop = {
    xApiClient: systemSheet.getRange(PROPERTY_CELL.X_CLIENT_KEY).getValue(),
    xApiClientSecret: systemSheet.getRange(PROPERTY_CELL.X_CLIENT_SECRET).getValue(),
    xOauth2Twitter: systemSheet.getRange(PROPERTY_CELL.X_OAUTH2_TWITTER).getValue(),
    xCodeVerifier: systemSheet.getRange(PROPERTY_CELL.X_CODE_VERIFIER).getValue(),
    xUserId: systemSheet.getRange(PROPERTY_CELL.X_USER_ID).getValue()
  };
  
  return prop;
}

// 特定のプロパティを取得する関数
function getSystemPropertyValue(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const systemSheet = ss.getSheetByName(SHEETS_NAME.SYSTEM);
  return systemSheet.getRange(key).getValue();
}

// 特定のプロパティを設定する関数
function setSystemPropertyValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const systemSheet = ss.getSheetByName(SHEETS_NAME.SYSTEM);
  systemSheet.getRange(key).setValue(value);
  Logger.log('プロパティ設定しました。Key:' + key + '、Value:' + value);
}