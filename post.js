// post.gs
// 投稿処理に関する機能を提供するファイル

/**
 * 各SNSへの投稿を実行
 * @param {Sheet} sheet スプレッドシート
 * @param {array} rowData 行データ
 * @param {number} rowIndex 行番号
 * @param {array} tableData シートデータ
 * @param {string} sheetName 対象シート名
 */
function executePost(sheet, rowData, rowIndex, tableData, sheetName) {
  Logger.log(`行${rowIndex}の投稿処理を開始`);
  
  try {

    // 投稿内容を取得（実際の列は要調整）
    const content = rowData[CONFIG.POST_TEXT];
    Logger.log(`投稿内容を取得: ${content}`);

    // 数が多い場合に投稿が被ってしまうので、先に投稿日付を先に仮にいれる
    sheet.getRange(rowIndex, CONFIG.POST_STATUS_COL + CONFIG.SHEET_ARRAY_COL_DIF).setValue("投稿中");
    SpreadsheetApp.flush();

    // エラー用の変数を用意する
    let errorMessage = "";

    // 添付ファイル情報を取得する
    let attachmentFiles = getAttachmentImageMovies(sheet, rowIndex)

    // threadsの投稿を行う(チェックが入っていない場合はスキップ)
    let threadsPostInfo;
    
    if(rowData[CONFIG.POST_CHECK_THREADS]){
      threadsPostInfo = postThreads(rowData, rowIndex, attachmentFiles, tableData, sheetName);

      // IDが取得できない場合は投稿できない旨を出力するため、メッセージをセットする
      if(threadsPostInfo.id == ""){
        errorMessage = "Threadsへの投稿が失敗しました。"
      }

    }



    // Xへの添付ファイル情報を取得する
    // Xの投稿を行う(チェックが入っていない場合はスキップ)
    let xPostInfo;
    if(rowData[CONFIG.POST_CHECK_X]){
      let attachmentFilesForX = getAttachmentImageMoviesForX(sheet, rowIndex)
      xPostInfo = postX(rowData, rowIndex, attachmentFilesForX, tableData, sheetName);

      // IDが取得できない場合は投稿できない旨を出力するため、メッセージをセットする
      if(xPostInfo.tweetId == ""){
        if(errorMessage == ""){
          errorMessage = "Xへの投稿が失敗しました。"
        }else{
          errorMessage += "\nXへの投稿が失敗しました。"
        }
      }

    }


    // Instagramへの投稿を行う(チェックが入っていない場合はスキップ)
    let instaPostInfo;
    if(rowData[CONFIG.POST_CHECK_INSTA]){
      instaPostInfo = postInstagram(rowData, rowIndex, attachmentFiles);

      // IDが取得できない場合は投稿できない旨を出力するため、メッセージをセットする
      if(instaPostInfo.id == ""){
        if(errorMessage == ""){
          errorMessage = "Instagramへの投稿が失敗しました。"
        }else{
          errorMessage += "\nInstagramへの投稿が失敗しました。"
        }
      }

    }


    // 投稿成功時、状態を更新
    updatePostStatus(sheet, rowData, rowIndex, threadsPostInfo, xPostInfo, instaPostInfo, errorMessage);

    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log(`行${rowIndex}の投稿処理でエラー発生: ${e.message}`);
    throw e;
  }
}

/**
 * Xへの投稿を実行
 * @param {array} rowData 行データ
 * @param {number} rowIndex 行番号
 * @param {array} attachmentFiles 添付ファイルデータ
 * @param {array} tableData シートデータ
 * @param {string} sheetName 対象シート名
 */
function postX(rowData, rowIndex, attachmentFiles, tableData, sheetName){

  try{

    let resId = "";
    let quoteId = "";

    // 返信IDを取得する
    if(rowData[CONFIG.POST_RES] != ""){
      resId = searchUpwardsInValues(tableData, rowIndex, rowData[CONFIG.POST_RES], CONFIG.POST_RES, CONFIG.X_POST_ID_COL, CONFIG.START_ROW, sheetName);
    }

    // 引用元IDを取得する
    if(rowData[CONFIG.POST_QUOTE] != ""){
      quoteId = extractTweetId(rowData[CONFIG.POST_QUOTE]);
    }

    let postText = convertMentionInfo(rowData[CONFIG.POST_TEXT], SNS_CONVERT.COL.X);

    let xPostInfo = postTweetWithMultipleImages(postText, attachmentFiles, resId, quoteId);

    Logger.log(`行${rowIndex}: 投稿成功 (tweetId: ${xPostInfo.tweetId})`);
  
    return xPostInfo;
  }catch(ex)
  {
    return{
      tweetId: "",
      url: ""
    }
  }
}


/**
 * Threadsへの投稿を実行
 * @param {array} rowData 行データ
 * @param {number} rowIndex 行番号
 * @param {array} attachmentFiles 添付ファイルデータ
 * @param {array} tableData シートデータ
 * @param {string} sheetName 対象シート名
 */
function postThreads(rowData, rowIndex, attachmentFiles, tableData, sheetName){

  try{

    Logger.log(`postThreads:Start`);

    let creationId;
    
    // 返信IDを取得する
    let resId = "";
    if(rowData[CONFIG.POST_RES] != ""){
      resId = searchUpwardsInValues(tableData, rowIndex, rowData[CONFIG.POST_RES], CONFIG.POST_RES, CONFIG.POST_ID_COL, CONFIG.START_ROW, sheetName);
      Logger.log(`resId:${resId}`);
    }

    // 投稿できるように準備する
    let quoteId = "";
    if(rowData[CONFIG.POST_QUOTE_THREADS] != ""){
      quoteId = rowData[CONFIG.POST_QUOTE_THREADS]
    }

    let postText = convertMentionInfo(rowData[CONFIG.POST_TEXT], SNS_CONVERT.COL.THREADS);

    // 添付ファイルなし
    if(attachmentFiles.length == 0){
      creationId = singlePostTextOnly(postText, resId, quoteId);
    }
    // 添付ファイルが1つの場合
    else if(attachmentFiles.length == 1){
      // ファイルを添付する
      creationId = singlePostAttachFile(attachmentFiles[0].url, attachmentFiles[0].movieDirectUrl, attachmentFiles[0].fileCategory, postText, resId, quoteId);
      // if (attachmentFiles[0].fileCategory == CONFIG.STRING_IMAGE){
      //   // 画像のみ処理する　※動画はURLでしか指定できないため現時点で保留
      //   creationId = singlePostAttachFile(attachmentFiles[0].url, attachmentFiles[0].fileCategory, postText);
      // }
    }
    // 添付ファイルが2つ以上の場合
    else{
      let mediaIds = [];

      attachmentFiles.forEach((attachment, index) => {
        // ファイルを添付し、メディアIDを取得する
        mediaIds.push(uploadSingleImageVideo(attachment.url, attachment.movieDirectUrl, attachment.fileCategory, quoteId));

      });

      // カルーセルアルバムを作成する
      creationId = postCarouselContainer(mediaIds, postText);
    }

    // 添付したファイルを公開する
    let publish = puglishPostInfo(creationId);

    Logger.log(`行${rowIndex}: 投稿成功 (postId: ${publish.postId})`);
      
    // 投稿成功後、投稿した情報を取得する
    let threadsPostInfo = getThreadsPostInfoDetail(publish);
    return threadsPostInfo;
  
  }catch(ex){
    return {
      id: "",
      permalink: ""
    }
  }
}


/**
 * Instagramへの投稿を実行
 * @param {array} rowData 行データ
 * @param {number} rowIndex 行番号
 * @param {array} attachmentFiles 添付ファイルデータ
 */
function postInstagram(rowData, rowIndex, attachmentFiles){

  try{

    let creationId;
    if(attachmentFiles.length == 0){
      return {
        id: "",
        permalink: ""
      }
    }

    let postText = convertMentionInfo(rowData[CONFIG.POST_TEXT], SNS_CONVERT.COL.INSTAGRAM);

    // 添付ファイルが1つの場合
    if(attachmentFiles.length == 1){
      // ファイルを添付する
      creationId = instaSinglePostAttachFile(attachmentFiles[0].url, attachmentFiles[0].movieDirectUrl, attachmentFiles[0].fileCategory, CONFIG.ENUM_INSTA_UL_TYPE.SINGLE, postText);
    }
    // 添付ファイルが2つ以上の場合
    else{
      let mediaIds = [];

      attachmentFiles.forEach((attachment, index) => {
        // ファイルを添付し、メディアIDを取得する
        mediaIds.push(instaSinglePostAttachFile(attachment.url, attachment.movieDirectUrl, attachment.fileCategory, CONFIG.ENUM_INSTA_UL_TYPE.MULTI));

      });

      // カルーセルアルバムを作成する
    creationId = postInstaCarouselContainer(mediaIds, postText);
    }

    // 添付したファイルを公開する
    let publish = instaPublishPostInfo(creationId);

    Logger.log(`行${rowIndex}: 投稿成功 (postId: ${publish.id})`);
      
    // 投稿成功後、投稿した情報を取得する
    let instaPostInfo = getInstaPostInfos();

    return {
      id: instaPostInfo.business_discovery.media.data[0].id,
      permalink: instaPostInfo.business_discovery.media.data[0].permalink
    };
  }catch(ex){
    return {
      id: "",
      permalink: ""
    }
  }
}


/**
 * 投稿状態を更新
 * @param {Sheet} sheet スプレッドシート
 * @param {object} rowData 行データ
 * @param {number} row 行番号
 * @param {Object} threadsPostInfo Threads投稿メディア情報
 * @param {Object} xPostInfo X投稿メディア情報
 * @param {Object} instaPostInfo Instagram投稿メディア情報
 * @param {string} errorMessage エラーメッセージ
 */
function updatePostStatus(sheet, rowData, row, threadsPostInfo, xPostInfo, instaPostInfo, errorMessage) {
  Logger.log(`行${row}の投稿状態を更新開始`);
  
  try {
    // データ範囲を一括で取得
    const dataRange = sheet.getRange(row,
      CONFIG.POST_STATUS_COL + CONFIG.SHEET_ARRAY_COL_DIF,
      1,
      8 // O-V列を取得
    );

    // 現在のJST日時を取得
    const jstDate = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Tokyo" }));
    const formattedDate = Utilities.formatDate(
      jstDate,
      "Asia/Tokyo",
      'yyyy/MM/dd HH:mm:ss'
    );
    
    let xPostUrl = "";
    let xPostId = "";
    if(xPostInfo){
      xPostUrl = xPostInfo.url;
      xPostId = xPostInfo.tweetId;
    }

    let threadsPostUrl = "";
    let threadsPostId = "";
    if(threadsPostInfo){
      threadsPostUrl = threadsPostInfo.permalink;
      threadsPostId = threadsPostInfo.id;
    }

    let instaPostUrl = "";
    let instaPostId = "";
    if(instaPostInfo){
      instaPostUrl = instaPostInfo.permalink;
      instaPostId = instaPostInfo.id;
    }


    // 更新する値を配列で準備
    const updateValues = [[
      formattedDate,   // P列: 投稿日時
      xPostUrl,        // Q列: X投稿URL
      threadsPostUrl,  // R列: Threads投稿URL
      instaPostUrl,    // S列: Instagram投稿URL
      xPostId,         // T列: X投稿ID
      threadsPostId,   // U列: Threads投稿ID
      instaPostId,     // V列: Instagram投稿ID
      errorMessage     // W列: エラーメッセージ
    ]];
    
    // 値を一括で更新
    dataRange.setValues(updateValues);

    Logger.log(`行${row}の投稿状態を更新完了: ${formattedDate}`);
  } catch (e) {
    Logger.log(`行${row}の投稿状態更新でエラー: ${e.message}`);
    throw e;
  }
}



/**
 * Threadsへの投稿を実行
 * @param {string} content 投稿内容
 * @return {Object} 投稿結果
 */
function postToThreads(content) {
    Logger.log('Threads投稿処理を開始');
    
    // トークンの確認と更新
    if (!checkAndRefreshToken()) {
        const error = 'トークンの更新に失敗しました。再認証が必要です。';
        Logger.log(error);
        return {
            success: false,
            error: error
        };
    }

    // トークンデータの取得
    const tokenData = getThreadsToken();
    if (!tokenData) {
        const error = 'トークンの取得に失敗しました。認証を確認してください。';
        Logger.log(error);
        return {
            success: false,
            error: error
        };
    }
    Logger.log('トークン取得成功');

    const CREATE_URL = `https://graph.threads.net/v1.0/${tokenData.user_id}/threads`;
    const PUBLISH_URL = `https://graph.threads.net/v1.0/${tokenData.user_id}/threads_publish`;

    const createOptions = {
        method: "post",
        headers: {
            "Authorization": "Bearer " + tokenData.access_token,
            "Content-Type": "application/x-www-form-urlencoded"
        },
        payload: {
            text: content,
            media_type: "TEXT",
            access_token: tokenData.access_token
        }
    };

    try {
        Logger.log('スレッド作成リクエストを送信');
        const createResponse = UrlFetchApp.fetch(CREATE_URL, createOptions);
        Logger.log(`スレッド作成レスポンス: ${createResponse.getResponseCode()}`);
        
        Logger.log('API制限対策の待機開始（3秒）');
        Utilities.sleep(3000);
        Logger.log('待機完了');

        if (createResponse.getResponseCode() === 200) {
            const creationId = JSON.parse(createResponse.getContentText()).id;
            Logger.log(`スレッド作成成功: creationId=${creationId}`);
            
            const publishOptions = {
                method: "post",
                headers: {
                    "Authorization": "Bearer " + tokenData.access_token,
                    "Content-Type": "application/x-www-form-urlencoded"
                },
                payload: {
                    creation_id: creationId,
                    access_token: tokenData.access_token
                }
            };

            Logger.log('スレッド公開リクエストを送信');
            const publishResponse = UrlFetchApp.fetch(PUBLISH_URL, publishOptions);
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
        } else {
            const error = `スレッドの作成に失敗: ${createResponse.getResponseCode()} - ${createResponse.getContentText()}`;
            Logger.log(error);
            return {
                success: false,
                error: error
            };
        }
    } catch (error) {
        const errorMsg = `投稿処理でエラー発生: ${error.message}`;
        Logger.log(errorMsg);
        return {
            success: false,
            error: errorMsg
        };
    }
}

/**
 * トークンの状態を確認し、必要に応じて更新
 * @return {boolean} トークンが有効な場合はtrue
 */
function checkAndRefreshToken() {
    Logger.log('トークンの状態確認開始');
    const service = getServiceThreads();
    
    if (service.hasAccess()) {
        Logger.log('トークンは有効です');
        return true;
    } else {
        Logger.log('トークンが無効または期限切れです。更新を試みます。');
        try {
            if (service.refresh()) {
                Logger.log('トークンを更新しました');
                return true;
            } else {
                Logger.log('トークンの更新に失敗しました');
                return false;
            }
        } catch (e) {
            Logger.log(`トークン更新でエラー発生: ${e.toString()}`);
            return false;
        }
    }
}

/**
 * 投稿エラー時のリトライ処理
 * @param {function} postFunc 投稿処理を行う関数
 * @param {number} maxRetries 最大リトライ回数
 * @param {number} delay リトライ間隔（ミリ秒）
 * @return {Object} 投稿結果
 */
function retryPost(postFunc, maxRetries = 3, delay = 5000) {
    Logger.log(`リトライ処理を開始 (最大${maxRetries}回、間隔${delay}ms)`);
    
    for (let i = 0; i < maxRetries; i++) {
        Logger.log(`リトライ ${i + 1}/${maxRetries} 回目`);
        
        try {
            const result = postFunc();
            if (result.success) {
                Logger.log('リトライ成功');
                return result;
            }
            Logger.log(`リトライ失敗: ${result.error}`);
            
            if (i < maxRetries - 1) {
                Logger.log(`${delay}ms待機後、次のリトライを実行`);
                Utilities.sleep(delay);
            }
        } catch (error) {
            Logger.log(`リトライ中にエラー発生: ${error.message}`);
            if (i === maxRetries - 1) {
                const errorMsg = `最大リトライ回数到達: ${error.message}`;
                Logger.log(errorMsg);
                return {
                    success: false,
                    error: errorMsg
                };
            }
            Logger.log(`${delay}ms待機後、次のリトライを実行`);
            Utilities.sleep(delay);
        }
    }
    
    const errorMsg = '全てのリトライが失敗しました';
    Logger.log(errorMsg);
    return {
        success: false,
        error: errorMsg
    };
}