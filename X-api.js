// Twitter API関連URL
const LAMBDA_BASE = 'https://rfu5b5dar1.execute-api.ap-northeast-1.amazonaws.com/prod';
const INIT_PATH     = '/x-oauth/init';    // 認可開始用エンドポイント
const TOKEN_PATH    = '/x-oauth/token';   // トークン取得用エンドポイント

const INIT_URL      = LAMBDA_BASE + INIT_PATH;
const TOKEN_URL     = LAMBDA_BASE + TOKEN_PATH;

const ACCESS_TOKEN_URL = 'https://api.x.com/2/oauth2/token';


//********************************
//  認可開始：Lambda /init 呼び出し
//********************************
function authorizeLinkForNewTwitterBotWithImage() {
  const prop         = getSystemProperty();
  const clientId     = prop.xApiClient;
  const clientSecret = prop.xApiClientSecret;
  if (!clientId || !clientSecret) {
    throw new Error('CLIENT_ID または CLIENT_SECRET が設定されていません');
  }

  const url = INIT_URL
    + '?client_id='     + encodeURIComponent(clientId)
    + '&client_secret=' + encodeURIComponent(clientSecret);
  Logger.log('INIT_URL = ' + url);

  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  Logger.log('Response Code    = ' + resp.getResponseCode());
  Logger.log('Response Headers = ' + JSON.stringify(resp.getAllHeaders()));
  Logger.log('Response Body    = ' + resp.getContentText());

  if (resp.getResponseCode() !== 200) {
    throw new Error('認可開始エラー: ' + resp.getContentText());
  }

  const { authorizationUrl, state } = JSON.parse(resp.getContentText());
  if (!authorizationUrl || !state) {
    throw new Error('Lambda /init のレスポンスが不正です');
  }

  const sht = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SYSTEM_SHEET_NAME);
  if (!sht) {
    throw new Error(`シートが見つかりません: ${SYSTEM_SHEET_NAME}`);
  }
  sht.getRange(PROPERTY_CELL.X_STATE_CELL).setValue(state);

  const html = HtmlService
    .createHtmlOutput(
      `<p>下のリンクをクリックして X(Twitter) の認可を行ってください</p>
       <a href="${authorizationUrl}" target="_blank">${authorizationUrl}</a>`
    )
    .setWidth(600)
    .setHeight(160);

  SpreadsheetApp.getUi().showModalDialog(html, 'X アカウント認証');
}

//********************************
//  トークン取得：Lambda /token 呼び出し
//********************************
function fetchAccessToken() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sht   = ss.getSheetByName(SYSTEM_SHEET_NAME);
  const state = sht.getRange(PROPERTY_CELL.X_STATE_CELL).getValue();
  if (!state) {
    throw new Error('認可フロー未実行または state が見つかりません');
  }

  const url  = `${TOKEN_URL}?state=${encodeURIComponent(state)}`;
  Logger.log('>> TOKEN_URL = ' + url);

  const resp = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  Logger.log('>> Token Resp Code = ' + resp.getResponseCode());
  Logger.log('>> Token Resp Body = ' + resp.getContentText());

  if (resp.getResponseCode() !== 200) {
    throw new Error('トークン取得エラー: ' + resp.getContentText());
  }
  const data = JSON.parse(resp.getContentText());
  if (!data.access_token) {
    throw new Error('トークンレスポンス不正: ' + resp.getContentText());
  }

  // プロパティには必ず JSON 文字列で保存する
  data.timestamp = Date.now();
  setSystemProperty(
    PROPERTY_CELL.X_OAUTH2_TWITTER,
    JSON.stringify(data)
  );

  return data.access_token;
}



//********************************
//　ポスト情報送信用処理（トークン管理）
//********************************
function getXService() {
  const prop = getSystemProperty();
  let raw    = prop.xApiOauth2 || '';
  Logger.log('Existing tokenData(raw): ' + raw);

  // 不正な JSON ならクリア
  try {
    var tokenData = JSON.parse(raw);
  } catch (e) {
    Logger.log('❌ トークンデータが不正です。クリアします。');
    setSystemProperty(PROPERTY_CELL.X_OAUTH2_TWITTER, '');
    raw = '';
    tokenData = null;
  }

  if (tokenData) {
    const now    = Date.now();
    const ageSec = (now - tokenData.timestamp) / 1000;
    Logger.log(`Token age: ${ageSec}s / expires_in: ${tokenData.expires_in}s`);

    // 1) 有効期限内ならキャッシュそのまま
    if (tokenData.access_token && ageSec < tokenData.expires_in) {
      Logger.log('✔️ Using cached access_token');
      return {
        hasAccess:       () => true,
        getAccessToken:  () => tokenData.access_token
      };
    }

    // 2) リフレッシュトークンがあれば更新
    if (tokenData.refresh_token) {
      Logger.log('🔄 Attempting refresh with refresh_token');
      const newTokenData = refreshAccessToken(
        tokenData.refresh_token,
        prop.xApiClient,
        prop.xApiClientSecret
      );
      Logger.log('Refresh response data: ' + JSON.stringify(newTokenData));

      setSystemProperty(
        PROPERTY_CELL.X_OAUTH2_TWITTER,
        JSON.stringify(newTokenData)
      );

      return {
        hasAccess:      () => true,
        getAccessToken: () => newTokenData.access_token
      };
    }

    Logger.log('⚠️ Cached token expired and no refresh_token available');
  }

  // 3) 新規取得フェーズ
  Logger.log('➡️ Fetching new access_token from Lambda');
  const accessToken = fetchAccessToken();

  setSnsAccountSettingStatus(CONFIG.CELL_SETTING_STATUS_X);
  setSnsCheck(CONFIG.CELL_SETTING_CHECKBOX_X);

  return {
    hasAccess:      () => !!accessToken,
    getAccessToken: () => accessToken
  };
}



//********************************
// リフレッシュトークンで更新
//********************************
function refreshAccessToken(refreshToken, clientId, clientSecret) {
  const endpoint = ACCESS_TOKEN_URL;
  Logger.log('Refreshing token at: ' + endpoint);
  const payload = {
    grant_type:    'refresh_token',
    refresh_token: refreshToken
  };
  Logger.log('Refresh payload: ' + JSON.stringify(payload));

  const options = {
    method: 'post',
    payload: payload,
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
      'Content-Type':  'application/x-www-form-urlencoded'
    },
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(endpoint, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  Logger.log(`Refresh response code: ${code}, body: ${text}`);
  if (code !== 200) {
    throw new Error('リフレッシュ失敗: ' + text);
  }

  const tokenData = JSON.parse(text);
  tokenData.timestamp = Date.now();
  return tokenData;
}



//********************************
// 認証コールバック関数
//********************************
function authCallback(request) {
  const prop = getSystemProperty();

  Logger.log('コールバックリクエスト: ' + JSON.stringify(request.parameter));

  var service      = getXService();
  var codeVerifier = getSystemPropertyValue(PROPERTY_CELL.X_CODE_VERIFIER);
  var clientId     = prop.xApiClient;
  var clientSecret = prop.xApiClientSecret;
  var redirectUri  = getRedirectUri();

  var payload = {
    'code':          request.parameter.code,
    'code_verifier': codeVerifier,
    'grant_type':    'authorization_code',
    'redirect_uri':  redirectUri,
    'client_id':     clientId
  };

  var tokenOptions = {
    'method':         'post',
    'contentType':    'application/x-www-form-urlencoded',
    'payload':        payload,
    'headers': {
      'Authorization': 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret)
    },
    'muteHttpExceptions': true
  };

  try {
    var response     = UrlFetchApp.fetch(ACCESS_TOKEN_URL, tokenOptions);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log('トークンレスポンス コード: ' + responseCode);
    Logger.log('トークンレスポンス: ' + responseText);

    if (responseCode >= 200 && responseCode < 300) {
      var tokenData = JSON.parse(responseText);
      if (tokenData.access_token) {
        setSystemProperty(
          PROPERTY_CELL.X_OAUTH2_TWITTER,
          JSON.stringify({
            access_token:  tokenData.access_token,
            refresh_token: tokenData.refresh_token,
            expires_in:    tokenData.expires_in,
            timestamp:     Date.now()
          })
        );
        return HtmlService.createHtmlOutput(
          '<h3>認証が成功しました</h3>' +
          '<p>このタブを閉じて、スクリプトに戻ってください。</p>'
        );
      }
    }
    // エラー時 HTML
    return HtmlService.createHtmlOutput(
      `<h3>APIエラー</h3>
       <p>ステータスコード: ${responseCode}</p>
       <pre>${responseText}</pre>`
    );
  } catch (e) {
    Logger.log('トークン取得エラー: ' + e.toString());
    return HtmlService.createHtmlOutput(
      `<h3>リクエストエラー</h3>
       <p>${e.toString()}</p>`
    );
  }
}

/**
 * 認証情報からユーザーIDをセットする
 */
function setXUserId(){
  let userId = getUserIdFromApiKey();

  // システムプロパティを記録する
  setSystemProperty(PROPERTY_CELL.X_USER_ID, userId);  // XユーザーID
  
  const ui = SpreadsheetApp.getUi();
  ui.alert("ユーザーIDの設定が完了しました。");
  Logger.log('ユーザーIDの設定が完了しました。');
}


//********************************
// PKCE ヘルパー
//********************************
function generateCodeVerifier() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';
  let verifier = '';
  for (let i = 0; i < 128; i++) {
    verifier += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return verifier;
}

function generateCodeChallenge(verifier) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, verifier);
  const encoded = Utilities.base64Encode(rawHash);
  return encoded.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}


function getRedirectUri() {
  return 'https://script.google.com/macros/d/' +
         ScriptApp.getScriptId() +
         '/usercallback';
}



//********************************
// 認証URL取得 & リセット
//********************************
function getAuthorizationUrl() {
  resetAuth();

  // PKCE用 code_challenge
  const verifier       = generateCodeVerifier();
  const challenge      = generateCodeChallenge(verifier);
  setSystemProperty(PROPERTY_CELL.X_CODE_VERIFIER, verifier);
  Logger.log('code_verifier: ' + verifier);
  Logger.log('code_challenge: ' + challenge);

  const service        = getXService();
  service.setParam('code_challenge', challenge);
  const authUrl        = service.getAuthorizationUrl();
  Logger.log('認証URL: ' + authUrl);
  return authUrl;
}

function resetAuth() {
  setSystemProperty(PROPERTY_CELL.X_OAUTH2_TWITTER, '');
  setSystemProperty(PROPERTY_CELL.X_CODE_VERIFIER, '');
  Logger.log('認証状態をリセットしました。');
}


// 認証状態をリセット（トラブルシューティング用）
function resetAuth() {
  // ユーザープロパティから認証情報を削除
  setSystemProperty(PROPERTY_CELL.X_OAUTH2_TWITTER, "");
  setSystemProperty(PROPERTY_CELL.X_CODE_VERIFIER, "");
  
  Logger.log('認証状態をリセットしました。');
  return '認証状態をリセットしました。';
}

// 複数の画像をアップロードする関数
function uploadImagesForX(attachmentInfo, index) {
  // v2のメディアアップロードエンドポイントを使用
  const uploadMediaEndpoint = 'https://api.twitter.com/2/media/upload';
  
  // 認証情報を取得する
  const service = getXService();
  if (!service) {
    throw new Error("サービスの取得に失敗しました。");
  }

  // アクセストークンの確認
  const accessToken = service.getAccessToken();
  Logger.log('アクセストークン: ' + accessToken);

  Logger.log('X画像アップロード中:' + index + '件目...');
  Logger.log('=== 画像アップロード開始 ===');
  Logger.log(`インデックス: ${index}`);
  Logger.log(`URL: ${attachmentInfo.url}`);
  Logger.log(`FileID: ${attachmentInfo.fileId}`);

  try {
    let imageBlob;
    
    const file = DriveApp.getFileById(attachmentInfo.fileId);
    imageBlob = file.getBlob();
    Logger.log(`Drive File Name: ${file.getName()}`);
    Logger.log(`Drive File MimeType: ${file.getMimeType()}`);

    // 画像の詳細情報を出力
    Logger.log(`Blob ContentType: ${imageBlob.getContentType()}`);
    Logger.log(`Blob Name: ${imageBlob.getName()}`);
    Logger.log(`Blob Size: ${imageBlob.getBytes().length} bytes`);

    // MIMEタイプの検証
    const validMimeTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    const currentMimeType = imageBlob.getContentType();

    if (!validMimeTypes.includes(currentMimeType)) {
      Logger.log(`警告: 不適切なMIMEタイプ: ${currentMimeType}`);
      imageBlob.setContentType('image/jpeg');
      Logger.log(`MIMEタイプを image/jpeg に変更しました`);
    }

    // Blob オブジェクトをそのまま payload にセット
    var options = {
      method: "post",
      payload: { 
        media: imageBlob,
        media_category: CONFIG.X.MEDIA_CATEGORY.IMAGE
      },
      headers: { "Authorization": "Bearer " + accessToken },
      muteHttpExceptions: true
    };
    
    Logger.log(`Options: ${options}`);

    var response = UrlFetchApp.fetch(uploadMediaEndpoint, options);
    Logger.log("シンプルアップロード応答: " + response.getContentText());
    
    var result = JSON.parse(response.getContentText());
    Logger.log("レスポンス解析結果: " + JSON.stringify(result));
    
    if (result.errors) {
      throw new Error("Twitterシンプルアップロードエラー: " + JSON.stringify(result.errors));
    }
    
    var mediaId = result.data.id;
    if (!mediaId) {
      throw new Error("アップロード成功しましたが、メディアIDが取得できませんでした: " + response.getContentText());
    }
    
    Logger.log("取得メディアID: " + mediaId);
    return mediaId;

  } catch (error) {
    Logger.log(`X画像アップロード失敗: ${attachmentInfo.url}: ${error.toString()}`);
    throw error;
  }

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

    // if(attachVideoInfo.fileId != ""){
    //   // Google Driveから動画ファイルを取得
    //   videoFile = DriveApp.getFileById(attachVideoInfo.fileId);
    //   videoBlob = videoFile.getBlob();
    //   totalBytes = videoBlob.getBytes().length;
    //   mimeType = videoFile.getMimeType();

    // }else{
    videoInfo = getVideoInfo(attachVideoInfo.url);
    videoBlob = videoInfo.blob;
    totalBytes = videoInfo.totalBytes;
    mimeType = videoInfo.contentType;
    // }

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
    const mediaId = initializeVideoUpload(service, totalBytes, mimeType, CONFIG.X.MEDIA_CATEGORY.VIDEO);
    
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
function initializeVideoUpload(service, totalBytes, mimeType, mediaCategory) {
  const endpoint = 'https://api.twitter.com/2/media/upload';
  const payload = {
    command: 'INIT',
    total_bytes: JSON.stringify(totalBytes),
    media_type: mimeType,
    media_category: mediaCategory
  };
  
  const options = {
    method: 'POST',
    payload: payload,
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  };
  
  const response = UrlFetchApp.fetch(endpoint, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  Logger.log('初期化レスポンス:', responseText);
  
  if (responseCode !== 200 && responseCode !== 201 && responseCode !== 202) {
    throw new Error('動画アップロードの初期化に失敗しました: ' + responseText);
  }
  
  const responseData = JSON.parse(responseText);
  if (!responseData.data || !responseData.data.id) {
    throw new Error('media_idが取得できませんでした: ' + responseText);
  }
  
  Logger.log('動画アップロードの初期化成功 - media_id: ' + responseData.data.id);
  return responseData.data.id;
}

/**
 * 動画データのチャンク分割アップロード
 */
function appendVideoChunks(service, mediaId, videoBlob) {
  const chunkSize = 1024 * 1024; // 1MB chunks
  const totalBytes = videoBlob.getBytes().length;
  const chunks = Math.ceil(totalBytes / chunkSize);
  
  Logger.log('========== チャンクアップロード開始 ==========');
  Logger.log(`総バイト数: ${totalBytes} bytes (${Math.round(totalBytes / 1024 / 1024)}MB)`);
  Logger.log(`チャンクサイズ: ${chunkSize} bytes (${Math.round(chunkSize / 1024 / 1024)}MB)`);
  Logger.log(`総チャンク数: ${chunks}`);

  const bytes = videoBlob.getBytes();
  for (let i = 0; i < chunks; i++) {
    const start = i * chunkSize;
    const end = Math.min(start + chunkSize, totalBytes);
    
    Logger.log(`\n----- チャンク ${i + 1}/${chunks} アップロード開始 -----`);
    Logger.log(`チャンクサイズ: ${end - start} bytes`);

    try {
      const chunk = bytes.slice(start, end);

      const options = {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + service.getAccessToken(),
        },
        payload: {
          command: 'APPEND',
          media_id: mediaId,
          segment_index: JSON.stringify(i),
          media: Utilities.newBlob(chunk, "application/octet-stream", "chunk" + i),
        },
        muteHttpExceptions: true
      };

      // optionsをログ出力
      Logger.log('=== APPEND Payload ===');
      Logger.log('command: ' + options.payload.command);
      Logger.log('media_id: ' + options.payload.media_id);
      Logger.log('segment_index: ' + options.payload.segment_index);
      Logger.log('media: ' + options.payload.media);
      Logger.log('===================');

      const response = UrlFetchApp.fetch('https://api.twitter.com/2/media/upload', options);
      const responseCode = response.getResponseCode();
      
      if (responseCode !== 200 && responseCode !== 201 && responseCode !== 202 && responseCode !== 204) {
        const responseText = response.getContentText();
        throw new Error(
          `チャンクアップロード失敗:\n` +
          `ステータスコード: ${responseCode}\n` +
          `レスポンス: ${responseText}`
        );
      }

      Logger.log(`チャンク ${i + 1}/${chunks} アップロード成功 (ステータスコード: ${responseCode})`);

      // メモリを解放
      delete chunk;

      if (i > 0 && i % 5 === 0) {
        Logger.log('チェックポイント - 短い休止を入れます');
        Utilities.sleep(2000);
      }

    } catch (error) {
      Logger.log('\n===== エラー詳細 =====');
      Logger.log(error.toString());
      throw new Error(`チャンク ${i + 1}/${chunks} のアップロードに失敗しました: ${error.toString()}`);
    }

    Utilities.sleep(1000);
  }

  delete bytes;
  Logger.log('\n========== 全チャンクのアップロード完了 ==========');
}

/**
 * 動画アップロードの完了通知
 */
function finalizeVideoUpload(service, mediaId) {
  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: {
      command: 'FINALIZE',
      media_id: mediaId
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://api.twitter.com/2/media/upload', options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200 && responseCode !== 201 && responseCode !== 202) {
    throw new Error('動画アップロードの完了処理に失敗しました: ' + response.getContentText());
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * 動画処理の完了待機
 */
function waitForVideoProcessing(service, mediaId) {
  const maxAttempts = 30;
  let attempts = 0;
  
  while (attempts < maxAttempts) {
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken()
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(
      'https://api.twitter.com/2/media/upload?command=STATUS&media_id=' + mediaId,
      options
    );
    
    if (response.getResponseCode() !== 200) {
      throw new Error('ステータス確認に失敗しました: ' + response.getContentText());
    }
    
    const status = JSON.parse(response.getContentText());
    
    if (status.data && status.data.processing_info) {
      if (status.data.processing_info.state === 'succeeded') {
        return true;
      } else if (status.data.processing_info.state === 'failed') {
        throw new Error('動画の処理に失敗しました: ' + JSON.stringify(status.data.processing_info.error));
      }
    }
    
    Logger.log(`動画処理待機中... 試行回数: ${attempts + 1}/${maxAttempts}`);
    attempts++;
    Utilities.sleep(2000);
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

      // 添付ファイルの情報をログ出力
      Logger.log('attachmentInfo: ' + JSON.stringify(attachmentInfo));

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

    // // 添付ファイルがあれば添付する
    // if (mediaIds.length > 0) {
    //   payloadObj.media = { 
    //     media_keys: mediaIds  // media_idsからmedia_keysに変更
    //   };
    // }
    // 添付ファイルがあればmediaプロパティとして追記
    if (mediaIds.length > 0) {
      payloadObj.media = {
        media_ids: mediaIds
      };
    }

    const options = {
      method: "post",
      payload: JSON.stringify(payloadObj),
      contentType: "application/json",
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken()
      }
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
 * @param {string} mediaCategory メディアカテゴリ
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
      media_category: mediaCategory,
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken()
      }
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
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken()
    }
  };

  const response = UrlFetchApp.fetch(url, options);
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
  Logger.log('--- POST Tweet リクエスト前 ---');
  Logger.log('URL: ' + url);
  Logger.log('Headers: ' + JSON.stringify(options.headers));
  Logger.log('Payload: ' + options.payload);


  const response = UrlFetchApp.fetch(url, options);
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
  setSystemProperty(PROPERTY_CELL.X_OAUTH2_TWITTER, "");
  setSystemProperty(PROPERTY_CELL.X_CODE_VERIFIER, "");
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