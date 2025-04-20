// const THREADS_CLIENT_ID = '1587515402190570';
// const THREADS_CLIENT_SECRET = '0b28b0575aec454f64288e80213646d9';
const THREADS_SCOPE = 'threads_basic,threads_content_publish';
const THREADS_AUTH_URL = 'https://threads.net/oauth/authorize';
const THREADS_TOKEN_URL = 'https://graph.threads.net/oauth/access_token';
// const THREADS_SPREADSHEET_ID = '1pZ-yZ6wRGCAiSPf3ayATbVE6Jf0nP3OOOKBHMLuxtpg';

function getServiceThreads() {

  // システムプロパティを取得
  const prop = getSystemProperty();

  return OAuth2.createService('threads')
      .setAuthorizationBaseUrl(THREADS_AUTH_URL)
      .setTokenUrl(THREADS_TOKEN_URL)
      .setClientId(prop.threadsClientId)
      .setClientSecret(prop.threadsClientSecret)
      .setCallbackFunction('authCallbackThreads')
      .setPropertyStore(PropertiesService.getScriptProperties())  // UserProperties から ScriptProperties に変更
      .setScope(THREADS_SCOPE)
      .setParam('access_type', 'offline')
      .setParam('prompt', 'consent')
      .setTokenHeaders({
          'Authorization': 'Basic ' + Utilities.base64Encode(prop.threadsClientId + ':' + prop.threadsClientSecret)
      });
}

function exchangeForLongTermToken(shortTermToken) {
    Logger.log('短期トークン：' + shortTermToken);

  // システムプロパティを取得
  const prop = getSystemProperty();

    const url = 'https://graph.threads.net/access_token' +
                '?grant_type=th_exchange_token' +
                '&client_secret=' + prop.threadsClientSecret +
                '&access_token=' + shortTermToken;

    Logger.log('URL：' + url);

    try {
        const response = UrlFetchApp.fetch(url);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        Logger.log('レスポンスコード：' + responseCode);
        Logger.log('レスポンス内容：' + responseText);

        if (responseCode === 200) {
            const responseBody = JSON.parse(responseText);
            const longTermToken = responseBody.access_token;
            const expiresIn = responseBody.expires_in || 5184000;
            
            Logger.log('取得した長期トークン：' + longTermToken);
            Logger.log('有効期限（秒）：' + expiresIn);

            // システムプロパティを記録する
            setSystemProperty(PROPERTY_CELL.THREADS_LONG_TIME_TOKEN, longTermToken);  // 長期トークン
            setSnsAccountSettingStatus(CONFIG.CELL_SETTING_STATUS_THREADS); // ステータス：設定済
            setSnsCheck(CONFIG.CELL_SETTING_CHECKBOX_THREADS);  // チェックボックスオン
            
            Logger.log('長期トークンを取得・保存しました');
            return longTermToken;
        } else {
            Logger.log('長期トークン取得エラー: ' + responseText);
            return null;
        }
    } catch (error) {
        Logger.log('トークン交換エラー: ' + error.toString());
        Logger.log('エラーの詳細: ' + JSON.stringify(error));
        return null;
    }
}


function authCallbackThreads(request){
    Logger.log('認証コールバック開始');
    Logger.log('リクエスト内容：' + JSON.stringify(request));

    var service = getServiceThreads();
    var isAuthorized = service.handleCallback(request);
    
    Logger.log('認証状態：' + isAuthorized);

    if (isAuthorized) {
        const shortTermToken = service.getAccessToken();
        Logger.log('取得した短期トークン：' + shortTermToken);
        
        const longTermToken = exchangeForLongTermToken(shortTermToken);
        Logger.log('長期トークン交換結果：' + (longTermToken ? '成功' : '失敗'));
        
        if (longTermToken) {
          // アカウントステータスを更新する
          setSnsAccountSettingStatus(CONFIG.CELL_SETTING_STATUS_THREADS);

          return HtmlService.createHtmlOutput('Threads認証成功！長期トークンの取得にも成功しました。このウィンドウを閉じて、スクリプトエディタに戻ってください。');
        } else {
          return HtmlService.createHtmlOutput('Threads認証は成功しましたが、長期トークンの取得に失敗しました。もう一度お試しください。');
        }
    } else {
        Logger.log('認証失敗');
        return HtmlService.createHtmlOutput('Threads認証に失敗しました。もう一度お試しください。');
    }
}

function getThreadsToken() {

  // プロパティを取得
  const prop = getSystemProperty();
  
  // const scriptProperties = PropertiesService.getScriptProperties();
  // const longTermToken = scriptProperties.getProperty('threads_long_term_token');
  // const tokenExpiry = scriptProperties.getProperty('threads_token_expiry');

  // 長期トークンが存在し、まだ有効な場合はそれを使用
  if (prop.threadsLongTimeToken) {
      try {
          const response = UrlFetchApp.fetch(
              'https://graph.threads.net/me?fields=id',
              {
                  headers: {
                      Authorization: 'Bearer ' + prop.threadsLongTimeToken
                  }
              }
          );
          const userId = JSON.parse(response.getContentText()).id;
          return {
              access_token: prop.threadsLongTimeToken,
              user_id: userId
          };
      } catch (error) {
          Logger.log('保存された長期トークンが無効です: ' + error.toString());
      }
  }

  // 長期トークンが無効な場合は通常のフローで取得
  var service = getServiceThreads();
  if (service.hasAccess()) {
      var accessToken = service.getAccessToken();
      // 短期トークンを長期トークンに交換
      const longTermToken = exchangeForLongTermToken(accessToken);
      
      if (longTermToken) {
          const response = UrlFetchApp.fetch(
              'https://graph.threads.net/me?fields=id',
              {
                  headers: {
                      Authorization: 'Bearer ' + longTermToken
                  }
              }
          );
          const userId = JSON.parse(response.getContentText()).id;
          return {
              access_token: longTermToken,
              user_id: userId
          };
      }
  }
  
  Logger.log('認証が必要です。getAuthorizationUrl()を実行してください。');
  return null;
}


function refreshLongTermToken() {
  
    // プロパティを取得
  const prop = getSystemProperty();

// const scriptProperties = PropertiesService.getScriptProperties();
  // const longTermToken = scriptProperties.getProperty('threads_long_term_token');

  const url = 'https://graph.threads.net/refresh_access_token' +
              '?grant_type=th_refresh_token' +
              '&access_token=' + prop.threadsLongTimeToken;

  try {
      const response = UrlFetchApp.fetch(url);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log('レスポンスコード：' + responseCode);
      Logger.log('レスポンス内容：' + responseText);

      if (responseCode === 200) {
          const responseBody = JSON.parse(responseText);
          const longTermToken = responseBody.access_token;
          const expiresIn = responseBody.expires_in || 5184000;
          
          Logger.log('取得した長期トークン：' + longTermToken);
          Logger.log('有効期限（秒）：' + expiresIn);

          // システムプロパティを記録する
          setSystemProperty(PROPERTY_CELL.THREADS_LONG_TIME_TOKEN, longTermToken);  // 長期トークン

          // const scriptProperties = PropertiesService.getScriptProperties();
          // scriptProperties.setProperty('threads_long_term_token', longTermToken);
          // scriptProperties.setProperty('threads_token_expiry', new Date(Date.now() + (expiresIn * 1000)).toISOString());
          
          Logger.log('長期トークンを取得・保存しました');
          return longTermToken;
      } else {
          Logger.log('長期トークン取得エラー: ' + responseText);
          return null;
      }
  } catch (error) {
      Logger.log('トークン交換エラー: ' + error.toString());
      Logger.log('エラーの詳細: ' + JSON.stringify(error));
      return null;
  }
}

function clearStoredTokens() {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('oauth2.threads');
    setSystemProperty(PROPERTY_CELL.THREADS_LONG_TIME_TOKEN, "");  // 長期トークン
    clearSnsAccountSettingStatus(CONFIG.CELL_SETTING_STATUS_THREADS); // ステータス：未設定
    clearSnsCheck(CONFIG.CELL_SETTING_CHECKBOX_THREADS);  // チェックボックスオフ
    
    Logger.log('すべての保存されたトークンをクリアしました。');
}

function getThreadsAuthorizationUrl() {
    var service = getServiceThreads();
    if (!service.hasAccess()) {
        var authorizationUrl = service.getAuthorizationUrl();
        Logger.log('以下のURLにアクセスして認証を行ってください:');
        Logger.log(authorizationUrl);
        return authorizationUrl;
    } else {
        Logger.log('すでに認証されています。');
        return null;
    }
}

/**
 * Threadsの投稿の情報を取得
 * @param {object} postInfo ポスト情報
 */
function getThreadsPostInfoDetail(postInfo){

  Logger.log('メディアID：' + postInfo.postId );

  // プロパティを取得
  const prop = getSystemProperty();

  // const scriptProperties = PropertiesService.getScriptProperties();
  // const longTermToken = scriptProperties.getProperty('threads_long_term_token');

  // メディア情報を取得する
  const url = 'https://graph.threads.net/v1.0/' + postInfo.postId +
              '?fields=id,media_product_type,media_type,media_url,permalink,owner,username,text,timestamp,shortcode,thumbnail_url,children,is_quote_post' +
              '&access_token=' + prop.threadsLongTimeToken;

  Logger.log('URL：' + url);

  try {
      const response = UrlFetchApp.fetch(url);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log('レスポンスコード：' + responseCode);
      Logger.log('レスポンス内容：' + responseText);

      if (responseCode === 200) {
          const responseBody = JSON.parse(responseText);
          return responseBody;
      } else {
          Logger.log('メディアID取得エラー: ' + responseText);
          return null;
      }
  } catch (error) {
      Logger.log('メディアID取得エラー: ' + error.toString());
      Logger.log('エラーの詳細: ' + JSON.stringify(error));
      return null;
  }

}


/**
 * テキストのみ投稿を作成する
 * @param {string} text テキスト
 * @param {string} resId 返信元ID
 * @param {string} quotePostId 引用元ID
 */
function singlePostTextOnly(text, resId, quotePostId){

  Logger.log('singlePostTextOnly:Start');
  Logger.log('resId:' + resId);

  // プロパティを取得
  const prop = getSystemProperty();

  // // アクセストークン取得
  // const scriptProperties = PropertiesService.getScriptProperties();
  // const longTermToken = scriptProperties.getProperty('threads_long_term_token');

  // ユーザーIDを取得する
  let userinfo = getThreadsToken();

  // URL定義
  let url = 'https://graph.threads.net/v1.0/' + userinfo.user_id + '/threads';

  // 基本的なpayloadの設定
  let payload = {
    text: text,
    media_type: "TEXT",
    access_token: prop.threadsLongTimeToken
  };

  // resIdが存在する場合のみ、payloadに追加
  if (resId != "") {
    payload.reply_to_id = resId;
  }

  // quotePostIdが存在する場合のみ、payloadに追加
  if (quotePostId != "") {
    payload.quote_post_id = quotePostId;
  }

  const createOptions = {
      method: "post",
      headers: {
          "Authorization": "Bearer " + prop.threadsLongTimeToken,
          "Content-Type": "application/x-www-form-urlencoded"
      },
      payload: payload
  };

  try {
    Logger.log('スレッド作成リクエストを送信');
    const response = UrlFetchApp.fetch(url, createOptions);
    Logger.log(`スレッド作成レスポンス: ${response.getResponseCode()}`);
    
    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('レスポンスコード：' + responseCode);
    Logger.log('レスポンス内容：' + responseText);

    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    if (responseCode === 200) {
        const creationId = JSON.parse(response.getContentText()).id;
        Logger.log(`スレッド作成成功: creationId=${creationId}`);
          
        return creationId;
    } else {
        Logger.log('スレッド作成エラー: ' + responseText);
        return null;
    }
  } 
  catch (error) {
      Logger.log('スレッド作成エラー: ' + error.toString());
      Logger.log('エラーの詳細: ' + JSON.stringify(error));
      return null;
  }

}


/**
 * 1画像/動画の投稿を作成する
 * @param {string} fileUrl ファイルURL
 * @param {string} movieUrl 動画ファイルURL
 * @param {string} fileType ファイルタイプ(IMAGE/VIDEO)
 * @param {string} text テキスト
 * @param {string} resId 返信元ID
 * @param {string} quotePostId 引用元ID
 */
function singlePostAttachFile(fileUrl, movieUrl, fileType, text, resId, quotePostId){

  Logger.log('singlePostAttachFile:Start');
  Logger.log('fileUrl:' + fileUrl);
  Logger.log('movieUrl:' + movieUrl);
  Logger.log('fileType:' + fileType);
  Logger.log('text:' + text);
  Logger.log('resId:' + resId);

  // プロパティを取得
  const prop = getSystemProperty();


  // アクセストークン取得
  // const scriptProperties = PropertiesService.getScriptProperties();
  // const longTermToken = scriptProperties.getProperty('threads_long_term_token');

  // ユーザーIDを取得する
  let userinfo = getThreadsToken();

  // URL定義
  let url = 'https://graph.threads.net/v1.0/' + userinfo.user_id + '/threads';


  // 基本的なpayloadの設定
  let payload = {
    text: text,
    access_token: prop.threadsLongTimeToken
  };

  // resIdが存在する場合のみ、payloadに追加
  if (resId != "") {
    payload.reply_to_id = resId;
  }

  // quotePostIdが存在する場合のみ、payloadに追加
  if (quotePostId != "") {
    payload.quote_post_id = quotePostId;
  }

  if(fileType == "image"){
    payload.media_type = "IMAGE";
    payload.image_url = fileUrl
  }else{
    payload.media_type = "VIDEO";
    payload.video_url = movieUrl;
  }

  let createOptions = {
    method: "post",
    headers: {
        "Authorization": "Bearer " + prop.threadsLongTimeToken,
        "Content-Type": "application/x-www-form-urlencoded"
    },
    payload: payload
  };


  try {
    Logger.log('スレッド作成リクエストを送信');
    //const response = UrlFetchApp.fetch(url);
    const response = UrlFetchApp.fetch(url, createOptions);
    Logger.log(`スレッド作成レスポンス: ${response.getResponseCode()}`);
    
    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('レスポンスコード：' + responseCode);
    Logger.log('レスポンス内容：' + responseText);

    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    if (responseCode === 200) {
        const creationId = JSON.parse(response.getContentText()).id;
        Logger.log(`スレッド作成成功: creationId=${creationId}`);
          
        return creationId;
    } else {
        Logger.log('スレッド作成エラー: ' + responseText);
        return null;
    }
  } 
  catch (error) {
      Logger.log('スレッド作成エラー: ' + error.toString());
      Logger.log('エラーの詳細: ' + JSON.stringify(error));
      return null;
  }

}



/**
 * Threadsメディアコンテナを作成する
 * @param {string} fileUrl ファイルURL
 * @param {string} movieUrl 動画ファイルURL
 * @param {string} fileType ファイルタイプ(image/video)
 * @param {string} resId 返信元ID
 * @param {string} quotePostId 引用元ID
 */
function uploadSingleImageVideo(fileUrl, movieUrl, fileType, resId, quotePostId){

  Logger.log('uploadSingleImageVideo:Start');
  Logger.log('fileUrl：' + fileUrl);
  Logger.log('movieUrl:' + movieUrl);
  Logger.log('fileType：' + fileType);
  Logger.log('resId:' + resId);

    // プロパティを取得
  const prop = getSystemProperty();

  // アクセストークン取得
  // const scriptProperties = PropertiesService.getScriptProperties();
  // const longTermToken = scriptProperties.getProperty('threads_long_term_token');

  // ユーザーIDを取得する
  let userinfo = getThreadsToken();

  // URL定義
  let url = 'https://graph.threads.net/v1.0/' + userinfo.user_id + '/threads';


  // 基本的なpayloadの設定
  let payload = {
    is_carousel_item: true,
    access_token: prop.threadsLongTimeToken
  };

  // resIdが存在する場合のみ、payloadに追加
  if (resId != "") {
    payload.reply_to_id = resId;
  }

  // quotePostIdが存在する場合のみ、payloadに追加
  if (quotePostId != "") {
    payload.quote_post_id = quotePostId;
  }

  if(fileType == "image"){
    payload.media_type = "IMAGE";
    payload.image_url = fileUrl
  }else{
    payload.media_type = "VIDEO";
    payload.video_url = movieUrl;
  }

  let createOptions = {
    method: "post",
    headers: {
        "Authorization": "Bearer " + prop.threadsLongTimeToken,
        "Content-Type": "application/x-www-form-urlencoded"
    },
    payload: payload
  };

  try {
    Logger.log('アイテムコンテナ作成リクエストを送信');
    const response = UrlFetchApp.fetch(url, createOptions);
    Logger.log(`アイテムコンテナ作成レスポンス: ${response.getResponseCode()}`);
    
    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('レスポンスコード：' + responseCode);
    Logger.log('レスポンス内容：' + responseText);

    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    if (responseCode === 200) {
        const mediaId = JSON.parse(response.getContentText()).id;
        Logger.log(`アイテムコンテナ作成成功: creationId=${mediaId}`);
          
        return mediaId;
    } else {
        Logger.log('アイテムコンテナ作成エラー: ' + responseText);
        return null;
    }
  } 
  catch (error) {
      Logger.log('アイテムコンテナ作成エラー: ' + error.toString());
      Logger.log('エラーの詳細: ' + JSON.stringify(error));
      return null;
  }


}


/**
 * カルーセルコンテナの作成
 * @param {array{number}} mediaIds メディアID配列
 * @param {string} text 投稿テキスト
  * @param {number} resId 返信元ID
 * @param {number} quotePostId 引用元ID
 */
function postCarouselContainer(mediaIds, text, resId, quotePostId){

  Logger.log('postCarouselContainer:Start');
  Logger.log('mediaIds:' + mediaIds);
  Logger.log('text:' + text);
  Logger.log('resId:' + resId);

  // プロパティを取得
  const prop = getSystemProperty();

  // ユーザーIDを取得する
  let userinfo = getThreadsToken();

  try {
    let retryLimitCount = 8;

    for(let tryCount = 0; tryCount<retryLimitCount; tryCount++){
      // 待ちフラグ
      let waitFlg = false;

      for(let i=0; i<mediaIds.length; i++){

        // コンテナステータスの確認（動画の場合に待ちが必要）
        let urlContainerStatus = 'https://graph.threads.net/v1.0/' + mediaIds[i] + `?fields=status,error_message&access_token=${prop.threadsLongTimeToken}`;
        Logger.log(`urlContainerStatus:${urlContainerStatus}`);

        let containerStatusResponse = UrlFetchApp.fetch(urlContainerStatus);
        let containerStatusJson = JSON.parse(containerStatusResponse.getContentText());

        // Containerステータスが完了以外の場合は待ちフラグをオンにして抜ける
        Logger.log(`コンテナステータス：${JSON.stringify(containerStatusJson)}}`);

        if(containerStatusJson.status != "FINISHED"){
          waitFlg = true;
          break;
        }
      }

      if(waitFlg){
        // 完了していない場合はスリープ
        Logger.log(`スリープ20秒`);
        Utilities.sleep(20000); 
      }else{
        break;
      }
    
    }


    // メディアIDをカンマ区切りでつなぐ
    let mediaIdString = mediaIds.join(",");

    // URL定義
    let url = 'https://graph.threads.net/v1.0/' + userinfo.user_id + '/threads';

    let payload = {
        text: text,
        media_type: "CAROUSEL",
        children: mediaIdString,
        access_token: prop.threadsLongTimeToken,
    }

    // resIdが存在する場合のみ、payloadに追加
    if (resId != "") {
      payload.reply_to_id = resId;
    }

    // quotePostIdが存在する場合のみ、payloadに追加
    if (quotePostId) {
      payload.quote_post_id = quotePostId;
    }


    let createOptions = {
      method: "post",
      headers: {
          "Authorization": "Bearer " + prop.threadsLongTimeToken,
          "Content-Type": "application/x-www-form-urlencoded"
      },
      payload: payload
    };

    Logger.log('カルーセルコンテナ作成リクエストを送信');
    const response = UrlFetchApp.fetch(url, createOptions);
    Logger.log(`カルーセルコンテナ作成レスポンス: ${response.getResponseCode()}`);
    
    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('レスポンスコード：' + responseCode);
    Logger.log('レスポンス内容：' + responseText);

    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

    if (responseCode === 200) {
        const creationId = JSON.parse(response.getContentText()).id;
        Logger.log(`カルーセルコンテナ作成成功: creationId=${creationId}`);
          
        return creationId;
    } else {
        Logger.log('カルーセルコンテナ作成エラー: ' + responseText);
        return null;
    }
  } 
  catch (error) {
      Logger.log('カルーセルコンテナ作成エラー: ' + error.toString());
      Logger.log('エラーの詳細: ' + JSON.stringify(error));
      return null;
  }

}


/**
 * Threads投稿を公開
 * @param {number} creationId 公開するメディアID
 */
function puglishPostInfo(creationId){

  Logger.log('puglishPostInfo:Start');
  Logger.log('id：' + creationId);

  // プロパティを取得
  const prop = getSystemProperty();

  // アクセストークン取得
  // const scriptProperties = PropertiesService.getScriptProperties();
  // const longTermToken = scriptProperties.getProperty('threads_long_term_token');

  // ユーザーIDを取得する
  let userinfo = getThreadsToken();

  // URL定義
  let url = 'https://graph.threads.net/v1.0/' + userinfo.user_id + '/threads_publish';

  // const publishOptions = {
  //   method: "post",
  //   headers: {
  //       "Authorization": "Bearer " + prop.threadsLongTimeToken,
  //       "Content-Type": "application/x-www-form-urlencoded"
  //   },
  //   payload: {
  //       creation_id: creationId,
  //       access_token: prop.threadsLongTimeToken
  //   }
  // };

  const payload = {
    creation_id: creationId,
    access_token: prop.threadsLongTimeToken
  };

  // payload を URLエンコード形式に変換
  const payloadString = Object.keys(payload).map(key => key + '=' + encodeURIComponent(payload[key])).join('&');

  const publishOptions = {
    method: "post",
    headers: {
        "Authorization": "Bearer " + prop.threadsLongTimeToken,
        "Content-Type": "application/x-www-form-urlencoded"
    },
    payload: payloadString
  };

  try {

    // コンテナステータスの確認（動画の場合に待ちが必要）
    let urlContainerStatus = 'https://graph.threads.net/v1.0/' + creationId + `?fields=status,error_message&access_token=${prop.threadsLongTimeToken}`;
    Logger.log(`urlContainerStatus:${urlContainerStatus}`);

    let retryLimitCount = 8;

    for(let tryCount = 0; tryCount<retryLimitCount; tryCount++){
      Logger.log(`tryCount:${tryCount}回目`);

      let containerStatusResponse = UrlFetchApp.fetch(urlContainerStatus);
      let containerStatusJson = JSON.parse(containerStatusResponse.getContentText());



      // Containerステータスが完了していたら抜ける
      Logger.log(`コンテナステータス：${JSON.stringify(containerStatusJson)}}`);
      // Logger.log(`コンテナステータス：${JSON.stringify(containerStatusResponse)}`);
      if(containerStatusJson.status == "FINISHED"){
        break;
      }
      
      // 完了していない場合はスリープ
      Logger.log(`スリープ20秒`);
      Utilities.sleep(20000); 
      
    }

    Logger.log('スレッド公開リクエストを送信');
    Logger.log(`url:${url}`);
    Logger.log(`publishOptions:${JSON.stringify(payload)}`);
    const publishResponse = UrlFetchApp.fetch(url, publishOptions);
    Logger.log(`スレッド公開レスポンス: ${publishResponse.getResponseCode()}`);

    if (publishResponse.getResponseCode() === 200) {
      const postId = JSON.parse(publishResponse.getContentText()).id;
      Logger.log(`投稿成功: postId=${postId}`);
      return {
          success: true,
          error: null,
          postId: postId
      };
    } else {
      const error = `投稿の公開に失敗: ${publishResponse.getResponseCode()} - ${publishResponse.getContentText()}`;
      Logger.log(error);
      return {
          success: false,
          error: error
      };
    }
  } 
  catch (error) {
      Logger.log('スレッド作成エラー: ' + error.toString());
      Logger.log('エラーの詳細: ' + JSON.stringify(error));
      return null;
  }

}


