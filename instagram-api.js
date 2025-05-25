const INSTAGRAM_API_VERSION = "v21.0";
const BASE_URL = 'https://graph.facebook.com';
/**
 * 認証情報からユーザーIDをセットする
 */
function setProperty(){

  const ui = SpreadsheetApp.getUi();
  //ui.alert("ユーザーIDの設定が完了しました。");
  Logger.log('ユーザーIDの設定が完了しました。');

}

/**
 * 長期アクセストークン発行
 */
function getLongTermAccessToken() {

  let prop = getSystemProperty();

  try {
    if (!prop.instaAppId) throw new Error("App IDが未入力です");
    if (!prop.instaAppSecret) throw new Error("App secretが未入力です");
    if (!prop.instaShortTimeToken) throw new Error("短期トークンが未入力です");
    const instaLongTermToken = generateLongTermAccessTokenParams(prop.instaAppId, prop.instaAppSecret, prop.instaShortTimeToken);

  } catch (err) {
    Logger.log('長期トークンの設定が失敗しました。');
    throw err;
  }
}

/***************************************
 * インスタグラム長期トークン取得
 * 短期トークン→長期トークンへの変換
 * *************************************/
function generateLongTermAccessTokenParams(appId, appSecret, shortTermAccessToken){

  try {
    Logger.log('短期トークン：' + shortTermAccessToken);

    const url = `https://graph.facebook.com/${INSTAGRAM_API_VERSION}/oauth/access_token` +
                '?grant_type=fb_exchange_token' +
                '&client_id=' + appId +
                '&client_secret=' + appSecret +
                '&fb_exchange_token=' + shortTermAccessToken;

    Logger.log('URL：' + url);

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
      setSystemProperty(PROPERTY_CELL.INSTA_LONG_ACCESS_TOKEN, longTermToken);  // 長期トークン

      // アカウントステータスを更新する
      setSnsAccountSettingStatus(CONFIG.CELL_SETTING_STATUS_INSTA); // ステータス：設定済
      setSnsCheck(CONFIG.CELL_SETTING_CHECKBOX_INSTA);  // チェックボックスオン

      Logger.log('長期トークンを取得・保存しました');

      // InstagramBusinessIDを取得
      // 画面から取得する　me/accounts?fields=instagram_business_account,name
      //getInstagramBusinessID(appId, appSecret, shortTermAccessToken);


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


/***************************************
 * 長期トークンのリフレッシュ
 ***************************************/
function getLongTokenRefresh(){

  // プロパティの取得
  let prop = getSystemProperty();

  try {
    Logger.log('長期トークン：' + prop.instaLongTimeToken);

    const url = `https://graph.instagram.com/refresh_access_token?grant_type=ig_refresh_token&access_token=${prop.instaLongTimeToken}`;

    Logger.log('URL：' + url);

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
      setSystemProperty(PROPERTY_CELL.INSTA_LONG_ACCESS_TOKEN, longTermToken);  // 長期トークン

      // const scriptProperties = PropertiesService.getScriptProperties();
      // scriptProperties.setProperty(PROPERTY_STRING.INSTA_LONG_ACCESS_TOKEN, longTermToken);
      // scriptProperties.setProperty(PROPERTY_STRING.INSTA_LONG_ACCESS_TOKEN_EXPIRY, new Date(Date.now() + (expiresIn * 1000)).toISOString());
      
      Logger.log('長期トークンを更新しました');
      return longTermToken;
    } else {
      Logger.log('長期トークン更新エラー: ' + responseText);
      return null;
    }
  } catch (error) {
    Logger.log('長期トークンを更新エラー: ' + error.toString());
    Logger.log('エラーの詳細: ' + JSON.stringify(error));
    return null;
  }
}

/**
 * [Instagram]
 * 1画像/動画の投稿を作成する
 * @param {string} fileUrl ファイルURL
 * @param {string} movieUrl 動画ファイルURL
 * @param {string} fileType ファイルタイプ(IMAGE/VIDEO)
 * @param {number} uploadType アップロードタイプ(1:単体、2:複数)
 * @param {string} text テキスト
 */
function instaSinglePostAttachFile(fileUrl, movieUrl, fileType, uploadType, text=""){

  Logger.log('instaSinglePostAttachFile:Start');
  Logger.log('fileUrl:' + fileUrl);
  Logger.log('fileType:' + fileType);
  Logger.log('uploadType:' + uploadType);
  Logger.log('text:' + text);

  // アクセストークン取得
  let prop = getSystemProperty();

  const endpoint = `${BASE_URL}/${INSTAGRAM_API_VERSION}/${prop.instaBusinessId}/media`;
  
  let payload;
  if(fileType == CONFIG.STRING_IMAGE){
     payload= {
      image_url: fileUrl,
      caption: text,
      access_token: prop.instaLongTimeToken
    };
  }else{
    if(uploadType ==  CONFIG.ENUM_INSTA_UL_TYPE.SINGLE){
      // videoかつ、リールの場合
      payload = {
        video_url: movieUrl,
        media_type: "REELS",
        caption: text,
        access_token: prop.instaLongTimeToken
      };
    }else{
      // videoかつ、カルーセル動画の場合
      payload = {
        video_url: movieUrl,
        media_type: "VIDEO",
        caption: text,
        is_carousel_item: true,
        access_token: prop.instaLongTimeToken
      };
    }

  }

  // Payloadをログに出力する
  Logger.log('payload:' + JSON.stringify(payload));

  try {
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      payload: payload
    });

    Logger.log('動画アップロード待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');


    if (response.getResponseCode() === 200) {
        const mediaId = JSON.parse(response.getContentText()).id;
        Logger.log(`カルーセルコンテナ作成成功: creationId=${mediaId}`);

        if(fileType == CONFIG.STRING_VIDEO){
          Logger.log('API制限対策の待機開始（2秒）');
          Utilities.sleep(2000);
          Logger.log('待機完了');
        }



        return mediaId;
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
 * [Instagram]
 * カルーセルコンテナの作成
 * @param {array{number}} mediaIds メディアID配列
 * @param {string} text 投稿テキスト
 */
function postInstaCarouselContainer(mediaIds, text){

  Logger.log('postInstaCarouselContainer:Start');
  Logger.log('mediaIds:' + mediaIds);
  Logger.log('text:' + text);


  // アクセストークン取得
  let prop = getSystemProperty();

  try {
    let retryLimitCount = 8;

    for(let tryCount = 0; tryCount<retryLimitCount; tryCount++){
      // 待ちフラグ
      let waitFlg = false;

      for(let i=0; i<mediaIds.length; i++){

        // コンテナステータスの確認（動画の場合に待ちが必要）
        let urlContainerStatus = `${BASE_URL}/${INSTAGRAM_API_VERSION}/${mediaIds[i]}/` + `?fields=status&access_token=${prop.instaLongTimeToken}`;
        Logger.log(`urlContainerStatus:${urlContainerStatus}`);

        let containerStatusResponse = UrlFetchApp.fetch(urlContainerStatus);
        let containerStatusJson = JSON.parse(containerStatusResponse.getContentText());

        // Containerステータスが完了以外の場合は待ちフラグをオンにして抜ける
        Logger.log(`コンテナステータス：${JSON.stringify(containerStatusJson)}}`);

      if(!containerStatusJson.status.includes("Finished")){
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
    const endpoint = `${BASE_URL}/${INSTAGRAM_API_VERSION}/${prop.instaBusinessId}/media`;

    const payload = {
      media_type: 'CAROUSEL',
      caption: text,
      children: mediaIdString,
      access_token: prop.instaLongTimeToken
    };

    Logger.log('カルーセルコンテナ作成リクエストを送信');
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      payload: payload
    });
    Logger.log(`カルーセルコンテナ作成レスポンス: ${response.getResponseCode()}`);

    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');


    if (response.getResponseCode() === 200) {
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
 * [Instagram]
 * 投稿を公開
 * @param {number} creationId 公開するメディアID
 */
function instaPublishPostInfo(creationId){

  Logger.log('publishPostInfo:Start');
  Logger.log('id：' + creationId);

  // アクセストークン取得
  let prop = getSystemProperty();
  
  try {

    // コンテナステータスの確認（動画の場合に待ちが必要）
    let urlContainerStatus = `${BASE_URL}/${INSTAGRAM_API_VERSION}/${creationId}/` + `?fields=status&access_token=${prop.instaLongTimeToken}`;
    //let urlContainerStatus = 'https://graph.instagram.com/' + creationId + `?fields=status&access_token=${prop.instaLongTimeToken}`;
    Logger.log(`urlContainerStatus:${urlContainerStatus}`);

    let retryLimitCount = 8;

    for(let tryCount = 0; tryCount<retryLimitCount; tryCount++){
      Logger.log(`tryCount:${tryCount}回目`);

      let containerStatusResponse = UrlFetchApp.fetch(urlContainerStatus);
      let containerStatusJson = JSON.parse(containerStatusResponse.getContentText());

      // Containerステータスが完了していたら抜ける
      Logger.log(`コンテナステータス：${JSON.stringify(containerStatusJson)}}`);
      // Logger.log(`コンテナステータス：${JSON.stringify(containerStatusResponse)}`);
      if(containerStatusJson.status.includes("Finished")){
        break;
      }
      
      // 完了していない場合はスリープ
      Logger.log(`スリープ20秒`);
      Utilities.sleep(20000); 
      
    }


    const endpoint = `${BASE_URL}/${INSTAGRAM_API_VERSION}/${prop.instaBusinessId}/media_publish`;
    
    const payload = {
      creation_id: creationId,
      access_token: prop.instaLongTimeToken
    };

    Logger.log('スレッド公開リクエストを送信');
    const publishResponse = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      payload: payload
    });
    Logger.log(`スレッド公開レスポンス: ${publishResponse.getResponseCode()}`);
    Logger.log('API制限対策の待機開始（2秒）');
    Utilities.sleep(2000);
    Logger.log('待機完了');

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
  } catch (error) {
    Logger.log('投稿の公開に失敗しました: ' + error);
    throw error;
  }
}


/**
 * [Instagram]
 * 投稿を取得する
 */
function getInstaPostInfos() {
  
  // プロパティ取得
  let prop = getSystemProperty();
  
  const endpoint = `https://graph.facebook.com/${INSTAGRAM_API_VERSION}/${prop.instaBusinessId}`;
  
  // メディアフィールドの構築
  let mediaField = 'media.limit(10)';
  mediaField += '{id,caption,media_type,permalink}';

  const fields = [
    'id',
    'username',
    mediaField
  ].join(',');

  const encodedFields = encodeURIComponent(`business_discovery.username(${prop.instaUserId}){${fields}}`);
  const url = `${endpoint}?fields=${encodedFields}&access_token=${prop.instaLongTimeToken}`;
  
  console.log(`Calling API with URL: ${url}`);

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    console.log(`API Response Code: ${responseCode}`);
    console.log(`API Response Body: ${responseBody}`);

    if (responseCode !== 200) {
      throw new Error(`API request failed with status ${responseCode}: ${responseBody}`);
    }
    
    const parsedResponse = JSON.parse(responseBody);
    
    // レスポンスから次のカーソルを取得（存在する場合）
    const mediaData = parsedResponse.business_discovery.media;
    if (mediaData && mediaData.paging && mediaData.paging.cursors) {
      parsedResponse.nextCursor = mediaData.paging.cursors.after;
    } else {
      parsedResponse.nextCursor = null;
    }
    
    return parsedResponse;
  } catch (error) {
    console.error(`Error in API call: ${error.message}`);
    throw error;
  }
}




/***************************************
 * コンテナIDを取得するメソッド
 ***************************************/
function makeInstaContenaAPI(imageUrl, fileCategory) {

  let prop = getSystemProperty();

  let postData;

  if(fileCategory == 'image'){
      postData = {
      image_url: imageUrl,
      media_type: '',
      is_carousel_item: true
    }
  }else{
      postData = {
      video_url: imageUrl,
      media_type: 'VIDEO',
      is_carousel_item: true
    }
  }
    
  const url = `https://graph.facebook.com/${INSTAGRAM_API_VERSION}/${prop.instaBusinessId}/media?`;
  const response = instagramApi(url, 'POST', postData);

  try {
    if (response) {
      const data = JSON.parse(response.getContentText());
      return data.id;
    } else {
      console.error('makeInstaContenaAPI:Error');
      return null;
    }
  } catch (error) {
    console.error('makeInstaContenaAPIError:', error);
    return null;
  }

}


/***************************************
 * グループ化コンテナIDを取得するメソッド
 ***************************************/
function makeInstaGroupContenaAPI(mediaIds, text) {

  Utilities.sleep(20000); //　DB登録を待つため一旦ストップ

  const postData = {
    media_type: 'CAROUSEL',
    caption: text,
    children: mediaIds
  }

  let prop = getSystemProperty();


  // グループコンテナID取得
  const url = `https://graph.facebook.com/${INSTAGRAM_API_VERSION}/${prop.instaBusinessId}/media?`;
  const response = instagramApi(url, 'POST', postData);

  try {
    if (response) {
      const data = JSON.parse(response.getContentText());
      return data.id;
    } else {
      console.error('makeInstaGroupContenaAPI: Error');
      return null;
    }
  } catch (error) {
    console.error('makeInstaGroupContenaAPIError:', error);
    return null;
  }

}


/***************************************
 * グループ化コンテナIDを公開
 ***************************************/
function puglishInstaPostInfo(creationId){

  // ステップ②関数（グループコンテナIDを作成）
  Utilities.sleep(20000); //　DB登録を待つため一旦ストップ

  // グループコンテナIDを使って投稿
  const contenaGroupId = creationId;

  const postData = {
    media_type: 'CAROUSEL',
    creation_id: contenaGroupId
  }

  let prop = getSystemProperty();

  const url = `https://graph.facebook.com/${INSTAGRAM_API_VERSION}/${prop.instaBusinessId}/media_publish?`;
  const response = instagramApi(url, 'POST', postData);

  try {
    if (response) {
      const data = JSON.parse(response.getContentText());
      return data;
    } else {
      console.error('puglishInstaPostInfo:Error');
      return null;
    }
  } catch (error) {
    console.error('puglishInstaPostInfo Error:', error);
    return null;
  }
}

/***************************************
 * APIを叩く関数
 * 短期トークン→長期トークンへの変換
 ***************************************/
function instagramApi(url, method,postData) {
 
  let prop = getSystemProperty();

  try {
    const data = postData
    const headers = {
      'Authorization': 'Bearer ' + prop.instaLongTimeToken,
      'Content-Type': 'application/json',
    };

    const options = {
      'method': method,
      'headers': headers,
      'payload': JSON.stringify(data)
    };

    const response = UrlFetchApp.fetch(url, options);
    return response;
  } catch (error) {
    console.error('Instagram APIのリクエスト中にエラーが発生しました:', error);
    return null;
  }
}


/**
 * サービスクリア（認証解除）
 */
function clearServiceForInstagram() {

  // アカウントステータスを更新する
  clearSnsAccountSettingStatus(CONFIG.CELL_SETTING_STATUS_INSTA); // ステータス：設定済
  clearSnsCheck(CONFIG.CELL_SETTING_CHECKBOX_INSTA);  // チェックボックスオン
}