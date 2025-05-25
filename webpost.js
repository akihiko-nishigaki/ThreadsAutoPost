/**
 * 投稿処理を統合管理するクラス
 */
class WebPost {
  constructor() {
    this.config = CONFIG;
  }

  /**
   * 投稿処理のメイン関数
   * @param {Object} data - 投稿データ
   * @returns {Object} 投稿結果
   */
  submitPost(data) {
    try {
      // プラットフォームデータの形式を変換
      const normalizedData = {
        ...data,
        platforms: {
          x: Array.isArray(data.platforms) ? data.platforms.includes('x') : data.platforms?.x || false,
          threads: Array.isArray(data.platforms) ? data.platforms.includes('threads') : data.platforms?.threads || false,
          instagram: Array.isArray(data.platforms) ? data.platforms.includes('instagram') : data.platforms?.instagram || false
        }
      };

      console.log('正規化されたデータ:', {
        text: normalizedData.text,
        files: normalizedData.files,
        platforms: normalizedData.platforms
      });

      // 入力値のバリデーション
      if (!this.validateInput(normalizedData)) {
        throw new Error('入力値が不正です');
      }

      // ファイルアップロード処理
      let fileUrls = null;
      if (normalizedData.files && normalizedData.files.length > 0) {
        fileUrls = this.uploadFiles(normalizedData.files);
      }

      // 各SNSへの投稿処理
      const results = {};
      
      if (normalizedData.platforms.x) {
        results.x = this.postToX(normalizedData.text, fileUrls);
      }
      
      if (normalizedData.platforms.threads) {
        results.threads = this.postToThreads(normalizedData.text, fileUrls);
      }
      
      if (normalizedData.platforms.instagram) {
        results.instagram = this.postToInstagram(normalizedData.text, fileUrls);
      }

      return {
        success: true,
        results: results
      };

    } catch (error) {
      console.error('投稿処理でエラーが発生しました:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * 入力値のバリデーション
   * @param {Object} data - 投稿データ
   * @returns {boolean} バリデーション結果
   */
  validateInput(data) {
    // テキストまたはファイルが必須
    if (!data.text && (!data.files || data.files.length === 0)) {
      console.error('バリデーションエラー: テキストまたはファイルが必須です');
      return false;
    }

    // プラットフォームのチェックを詳細化
    console.log('プラットフォームの状態:', {
      platformsObject: data.platforms,
      hasObject: !!data.platforms,
      x: data.platforms?.x,
      threads: data.platforms?.threads,
      instagram: data.platforms?.instagram
    });

    if (!data.platforms || 
        (!data.platforms.x && !data.platforms.threads && !data.platforms.instagram)) {
      console.error('バリデーションエラー: 少なくとも1つのプラットフォームを選択してください');
      return false;
    }

    // テキストの長さチェック
    if (data.text && data.text.length > this.config.APP.MAX_TEXT_LENGTH) {
      console.error('バリデーションエラー: テキストが長すぎます', {
        currentLength: data.text.length,
        maxLength: this.config.APP.MAX_TEXT_LENGTH
      });
      return false;
    }

    return true;
  }

  /**
   * ファイルアップロード処理
   * @param {Array} files - アップロードするファイルの配列
   * @returns {Object} アップロード結果
   */
  uploadFiles(files) {
    const results = [];
    
    for (const file of files) {
      try {
        const result = uploadFile(
          file.base64Data,
          file.fileName,
          file.mimeType
        );

        if (result.success) {
          results.push({
            url: result.url,
            type: file.mimeType.startsWith('image/') ? 'image' : 'video'
          });
        } else {
          throw new Error(`ファイルのアップロードに失敗しました: ${result.error}`);
        }
      } catch (error) {
        throw new Error(`ファイルのアップロード中にエラーが発生しました: ${error.message}`);
      }
    }

    return results;
  }

  /**
   * Xへの投稿処理
   * @param {string} text - 投稿テキスト
   * @param {Array} fileUrls - ファイルURLの配列
   * @returns {Object} 投稿結果
   */
  postToX(text, fileUrls) {
    try {
      if (fileUrls && fileUrls.length > 0) {
        // 画像または動画が添付されている場合
        if (fileUrls[0].type === 'image') {
          // 画像投稿
          return postTweetWithMultipleImages(text, fileUrls, "", "");
        } else {
          // 動画投稿
          return postTweetWithVideo(text, fileUrls[0].url);
        }
      } else {
        // テキストのみの投稿
        return postTweetWithMultipleImages(text, [], "", "");
      }
    } catch (error) {
      throw new Error(`Xへの投稿に失敗しました: ${error.message}`);
    }
  }

  /**
   * Threadsへの投稿処理
   * @param {string} text - 投稿テキスト
   * @param {Array} fileUrls - ファイルURLの配列
   * @returns {Object} 投稿結果
   */
  postToThreads(text, fileUrls) {
    try {
      let creationId;

      // トークンの取得を試みる
      const token = getThreadsToken();
      if (!token) {
        throw new Error('Threadsのトークン取得に失敗しました');
      }

      if (!fileUrls || fileUrls.length === 0) {
        // テキストのみの投稿
        console.log('テキストのみの投稿を試みます:', text);
        creationId = singlePostTextOnly(text, "", "");
        if (!creationId) {
          throw new Error('Threadsへのテキスト投稿に失敗しました');
        }
      } else if (fileUrls.length === 1) {
        // 単一ファイルの投稿
        const fileUrl = fileUrls[0].url;
        const movieUrl = fileUrls[0].type === 'video' ? fileUrl : null;
        const fileType = fileUrls[0].type === 'image' ? 'image' : 'video';
        
        console.log('単一ファイルの投稿を試みます:', { fileUrl, movieUrl, fileType, text });
        creationId = singlePostAttachFile(fileUrl, movieUrl, fileType, text, "", "");
        if (!creationId) {
          throw new Error('Threadsへのメディア投稿に失敗しました');
        }
      } else {
        // 複数ファイルの投稿（カルーセル）
        const mediaIds = [];
        for (const file of fileUrls) {
          const fileUrl = file.url;
          const movieUrl = file.type === 'video' ? fileUrl : null;
          const fileType = file.type === 'image' ? 'image' : 'video';
          
          console.log('メディアアップロードを試みます:', { fileUrl, movieUrl, fileType });
          const mediaId = uploadSingleImageVideo(fileUrl, movieUrl, fileType, "", "");
          if (!mediaId) {
            throw new Error('Threadsへのメディアアップロードに失敗しました');
          }
          mediaIds.push(mediaId);
        }

        console.log('カルーセル投稿を試みます:', { mediaIds, text });
        creationId = postCarouselContainer(mediaIds, text, "", "");
        if (!creationId) {
          throw new Error('Threadsへのカルーセル投稿に失敗しました');
        }
      }

      // 投稿を公開
      console.log('投稿の公開を試みます:', creationId);
      const publishResult = puglishPostInfo(creationId);
      if (!publishResult || !publishResult.success) {
        throw new Error(publishResult?.error || 'Threadsへの投稿公開に失敗しました');
      }

      return {
        success: true,
        postId: publishResult.postId
      };

    } catch (error) {
      console.error('Threads投稿エラーの詳細:', {
        error: error.message,
        stack: error.stack,
        text: text,
        fileUrls: fileUrls
      });
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Instagramへの投稿処理
   * @param {string} text - 投稿テキスト
   * @param {Array} fileUrls - ファイルURLの配列
   * @returns {Object} 投稿結果
   */
  postToInstagram(text, fileUrls) {
    try {
      if (!fileUrls || fileUrls.length === 0) {
        throw new Error('Instagramへの投稿には少なくとも1つのメディアファイルが必要です');
      }

      if (fileUrls.length === 1) {
        // 単一ファイルの投稿
        const fileUrl = fileUrls[0].url;
        const movieUrl = fileUrls[0].type === 'video' ? fileUrl : null;
        const fileType = fileUrls[0].type === 'image' ? CONFIG.STRING_IMAGE : CONFIG.STRING_VIDEO;
        const uploadType = CONFIG.ENUM_INSTA_UL_TYPE.SINGLE;
        
        return instaSinglePostAttachFile(fileUrl, movieUrl, fileType, uploadType, text);
      } else {
        // 複数ファイルの投稿（カルーセル）
        const mediaIds = fileUrls.map(file => 
          makeInstaContenaAPI(file.url, file.type === 'image' ? CONFIG.STRING_IMAGE : CONFIG.STRING_VIDEO)
        );
        return postInstaCarouselContainer(mediaIds, text);
      }
    } catch (error) {
      throw new Error(`Instagramへの投稿に失敗しました: ${error.message}`);
    }
  }
} 