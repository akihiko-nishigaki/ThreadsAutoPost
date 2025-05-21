// R2の認証情報
const REGION = 'auto';
const SERVICE = 's3';

/**
 * テスト用の関数：指定されたファイルIDを使用してR2へのアップロードをテストする
 * @param {string} fileId - Google DriveのファイルID
 */
function testupload() {
  uploadToCloudflareR2("16h0jr1RsN4o381jksCTWWn8-0kqoRfeC", "aiueo");
}

// ============ AWS4-HMAC-SHA256署名の生成に必要な補助関数群 ============

/**
 * HMAC-SHA256ハッシュを計算する
 * @param {Uint8Array|string} key - 暗号化キー
 * @param {string|Uint8Array} msg - ハッシュ化するメッセージ
 * @returns {Uint8Array} HMAC-SHA256ハッシュ値
 */
function hmacSHA256(key, msg) {
  var keyBlob = (key instanceof Uint8Array) ? key : Utilities.newBlob(key).getBytes();
  var msgBlob = (typeof msg === 'string') ? Utilities.newBlob(msg).getBytes() : msg;
  return Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, msgBlob, keyBlob);
}

/**
 * バイト配列を16進数文字列に変換する
 * @param {Uint8Array} bytes - 変換するバイト配列
 * @returns {string} 16進数文字列
 */
function bytesToHex(bytes) {
  return bytes.reduce(function(str, byte) {
    var hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return str + (hex.length === 1 ? '0' : '') + hex;
  }, '');
}

/**
 * SHA256ハッシュを計算し、16進数文字列として返す
 * @param {Uint8Array} bytes - ハッシュ化するバイト配列
 * @returns {string} SHA256ハッシュの16進数文字列
 */
function calculateSHA256(bytes) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return bytesToHex(digest);
}

/**
 * AWS4署名キーを生成する
 * @param {string} key - シークレットキー
 * @param {string} dateStamp - 日付（YYYYMMDD形式）
 * @param {string} regionName - リージョン名
 * @param {string} serviceName - サービス名
 * @returns {Uint8Array} 署名キー
 */
function getSignatureKey(key, dateStamp, regionName, serviceName) {
  var kSecret = Utilities.newBlob('AWS4' + key).getBytes();
  var kDate = hmacSHA256(kSecret, dateStamp);
  var kRegion = hmacSHA256(kDate, regionName);
  var kService = hmacSHA256(kRegion, serviceName);
  var kSigning = hmacSHA256(kService, 'aws4_request');
  return kSigning;
}

/**
 * Google DriveのファイルをCloudflare R2にアップロードする
 * @param {string} fileId - アップロードするファイルのGoogle DriveファイルID
 * @returns {string} アップロードされたファイルのURL
 */
function uploadToCloudflareR2(fileId, userId) {

  // システムプロパティを取得
  const prop = getSystemProperty();

  // ======== ファイルの準備 ========
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var bytes = blob.getBytes();
  var contentSha256 = calculateSHA256(bytes);
  
  // ======== 日付とタイムスタンプの生成 ========
  var now = new Date();
  var amzDate = now.toISOString().replace(/[:\-]|\.\d{3}/g, '');  // ISO8601形式から特殊文字を除去
  var dateStamp = amzDate.slice(0, 8);  // YYYYMMDD形式
  
  // ======== リクエストURLとパスの準備 ========
  var host = prop.cloudflareAccountId + '.r2.cloudflarestorage.com';
  var canonical_uri = '/' + prop.cloudflareBucket + '/' + encodeURIComponent(userId + '_' +file.getName());
  
  // ======== 正規リクエストの構築 ========
  // AWS署名バージョン4の仕様に従って正規リクエストを作成
  var canonical_headers = 'content-type:' + blob.getContentType() + '\n' +
                         'host:' + host + '\n' +
                         'x-amz-content-sha256:' + contentSha256 + '\n' +
                         'x-amz-date:' + amzDate + '\n';
  
  var signed_headers = 'content-type;host;x-amz-content-sha256;x-amz-date';
  
  var canonical_request = 'PUT\n' +
                         canonical_uri + '\n' +
                         '\n' + // 空のクエリ文字列
                         canonical_headers + '\n' +
                         signed_headers + '\n' +
                         contentSha256;
  
  // Logger.log('Canonical Request: ' + canonical_request);
  
  // ======== 署名の生成 ========
  var algorithm = 'AWS4-HMAC-SHA256';
  var credential_scope = dateStamp + '/' + REGION + '/' + SERVICE + '/aws4_request';
  
  // 署名対象の文字列を作成
  var string_to_sign = algorithm + '\n' +
                       amzDate + '\n' +
                       credential_scope + '\n' +
                       calculateSHA256(Utilities.newBlob(canonical_request).getBytes());
  
  // Logger.log('String to Sign: ' + string_to_sign);
  
  // 最終的な署名を計算
  var signing_key = getSignatureKey(prop.cloudflareSecretKey, dateStamp, REGION, SERVICE);
  var signature = bytesToHex(hmacSHA256(signing_key, string_to_sign));
  
  // ======== 認証ヘッダーの構築 ========
  var authorization_header = algorithm + ' ' +
                           'Credential=' + prop.cloudflareAccessKey + '/' + credential_scope + ',' +
                           'SignedHeaders=' + signed_headers + ',' +
                           'Signature=' + signature;
  
  // ======== リクエストの実行 ========
  var url = 'https://' + host + canonical_uri;
  
  var options = {
    'method': 'PUT',
    'headers': {
      'Authorization': authorization_header,
      'Content-Type': blob.getContentType(),
      'x-amz-content-sha256': contentSha256,
      'x-amz-date': amzDate
    },
    'payload': bytes,
    'muteHttpExceptions': true  // エラーレスポンスを取得するため
  };
  
  // デバッグ用のログ出力
  // Logger.log('Request URL: ' + url);
  // Logger.log('Request Headers: ' + JSON.stringify(options.headers, null, 2));
  
  // ======== アップロードの実行とレスポンスの処理 ========
  var response = UrlFetchApp.fetch(url, options);
  // Logger.log('Response Code: ' + response.getResponseCode());
  // Logger.log('Response Content: ' + response.getContentText());
  
  // レスポンスのチェックと結果の返却
  if (response.getResponseCode() === 200) {
    // パブリックアクセス用のURLを生成
    return prop.cloudflarePublicUrl + encodeURIComponent(userId + '_' + file.getName());
  } else {
    throw new Error('Upload failed: ' + response.getContentText());
  }
}


/**
 * R2バケット内の古いファイルを削除する定期実行用のメイン関数
 */
function cleanupOldFiles() {
  deleteOldFilesFromR2();
}

/**
 * R2バケットから1週間以上前のファイルを削除する
 */
function deleteOldFilesFromR2() {

  // システムプロパティを取得
  const prop = getSystemProperty();

  // リクエスト情報の準備
  var now = new Date();
  var amzDate = now.toISOString().replace(/[:\-]|\.\d{3}/g, '');
  var dateStamp = amzDate.slice(0, 8);
  var host = prop.cloudflareAccountId + '.r2.cloudflarestorage.com';

  // バケット内のオブジェクト一覧を取得するためのリクエストを準備
  var listRequest = generateSignedRequest('GET', host, prop.cloudflareBucket, '', {
    'list-type': '2'
  }, '', prop.cloudflareAccessKey, prop.cloudflareSecretKey, REGION, SERVICE, dateStamp, amzDate);

  try {
    // オブジェクト一覧を取得
    var response = UrlFetchApp.fetch(listRequest.url, listRequest.options);
    if (response.getResponseCode() !== 200) {
      throw new Error('Failed to list objects: ' + response.getContentText());
    }

    // XMLレスポンスをパース
    var xml = XmlService.parse(response.getContentText());
    var root = xml.getRootElement();
    var namespace = root.getNamespace();
    var contents = root.getChildren('Contents', namespace);

    // 1週間前の日時を計算
    var oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    // 各オブジェクトをチェックして古いものを削除
    contents.forEach(function(content) {
      var key = content.getChild('Key', namespace).getText();
      var lastModified = new Date(content.getChild('LastModified', namespace).getText());

      if (lastModified < oneWeekAgo) {
        // 削除リクエストを生成
        var deleteRequest = generateSignedRequest('DELETE', host, prop.cloudflareBucket, key, {}, '', 
          prop.cloudflareAccessKey, prop.cloudflareSecretKey, REGION, SERVICE, dateStamp, amzDate);

        try {
          var deleteResponse = UrlFetchApp.fetch(deleteRequest.url, deleteRequest.options);
          if (deleteResponse.getResponseCode() === 204) {
            Logger.log('Successfully deleted: ' + key);
          } else {
            Logger.log('Failed to delete ' + key + ': ' + deleteResponse.getContentText());
          }
        } catch (e) {
          Logger.log('Error deleting ' + key + ': ' + e.toString());
        }
      }
    });
  } catch (e) {
    Logger.log('Error in cleanup process: ' + e.toString());
    throw e;
  }
}

/**
 * AWS署名バージョン4に基づいて署名付きリクエストを生成する
 */
function generateSignedRequest(method, host, bucket, key, queryParams, payload,
                             accessKey, secretKey, region, service, dateStamp, amzDate) {
  var canonical_uri = '/' + bucket + (key ? '/' + encodeURIComponent(key) : '');
  
  // クエリパラメータを正規化
  var canonical_querystring = Object.keys(queryParams)
    .sort()
    .map(function(param) {
      return encodeURIComponent(param) + '=' + encodeURIComponent(queryParams[param]);
    })
    .join('&');

  var payloadHash = payload ? calculateSHA256(Utilities.newBlob(payload).getBytes()) : 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';

  // 正規リクエストの作成
  var canonical_headers = 'host:' + host + '\n' +
                         'x-amz-content-sha256:' + payloadHash + '\n' +
                         'x-amz-date:' + amzDate + '\n';
  
  var signed_headers = 'host;x-amz-content-sha256;x-amz-date';
  
  var canonical_request = method + '\n' +
                         canonical_uri + '\n' +
                         canonical_querystring + '\n' +
                         canonical_headers + '\n' +
                         signed_headers + '\n' +
                         payloadHash;

  // 署名の生成
  var algorithm = 'AWS4-HMAC-SHA256';
  var credential_scope = dateStamp + '/' + region + '/' + service + '/aws4_request';
  
  var string_to_sign = algorithm + '\n' +
                       amzDate + '\n' +
                       credential_scope + '\n' +
                       calculateSHA256(Utilities.newBlob(canonical_request).getBytes());
  
  var signing_key = getSignatureKey(secretKey, dateStamp, region, service);
  var signature = bytesToHex(hmacSHA256(signing_key, string_to_sign));
  
  // 認証ヘッダーの作成
  var authorization_header = algorithm + ' ' +
                           'Credential=' + accessKey + '/' + credential_scope + ',' +
                           'SignedHeaders=' + signed_headers + ',' +
                           'Signature=' + signature;

  var url = 'https://' + host + canonical_uri + (canonical_querystring ? '?' + canonical_querystring : '');
  
  var options = {
    'method': method,
    'headers': {
      'Authorization': authorization_header,
      'x-amz-content-sha256': payloadHash,
      'x-amz-date': amzDate
    },
    'muteHttpExceptions': true
  };

  if (payload) {
    options.payload = payload;
  }

  return {
    url: url,
    options: options
  };
}

/**
 * 削除実行前の確認用関数
 * 削除対象のファイルをリストアップするだけで、実際の削除は行わない
 */
function previewFilesToDelete() {
  // R2の認証情報
  const prop = getSystemProperty();
  
  var now = new Date();
  var amzDate = now.toISOString().replace(/[:\-]|\.\d{3}/g, '');
  var dateStamp = amzDate.slice(0, 8);
  var host = prop.cloudflareAccountId + '.r2.cloudflarestorage.com';

  var listRequest = generateSignedRequest('GET', host, prop.cloudflareBucket, '', {
    'list-type': '2'
  }, '', prop.cloudflareAccessKey, prop.cloudflareSecretKey, REGION, SERVICE, dateStamp, amzDate);

  try {
    var response = UrlFetchApp.fetch(listRequest.url, listRequest.options);
    if (response.getResponseCode() !== 200) {
      throw new Error('Failed to list objects: ' + response.getContentText());
    }

    var xml = XmlService.parse(response.getContentText());
    var root = xml.getRootElement();
    var namespace = root.getNamespace();
    var contents = root.getChildren('Contents', namespace);

    var oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    var filesToDelete = [];

    contents.forEach(function(content) {
      var key = content.getChild('Key', namespace).getText();
      var lastModified = new Date(content.getChild('LastModified', namespace).getText());
      var size = content.getChild('Size', namespace).getText();

      if (lastModified < oneWeekAgo) {
        filesToDelete.push({
          name: key,
          lastModified: lastModified.toISOString(),
          size: formatFileSize(parseInt(size))
        });
      }
    });

    Logger.log('以下のファイルが削除対象です：');
    filesToDelete.forEach(function(file) {
      Logger.log('ファイル名: ' + file.name);
      Logger.log('最終更新日: ' + file.lastModified);
      Logger.log('ファイルサイズ: ' + file.size);
      Logger.log('-------------------');
    });

    return filesToDelete;
  } catch (e) {
    Logger.log('プレビュー処理中にエラーが発生しました: ' + e.toString());
    throw e;
  }
}

/**
 * ファイルサイズを人間が読みやすい形式に変換する
 * @param {number} bytes - バイト数
 * @returns {string} 読みやすい形式のサイズ表示
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  var k = 1024;
  var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function cloudflareMovieUpload(fileId)
{
  let userInfo = Session.getActiveUser();
  let userName = getLocalPartFromEmail(userInfo.getEmail());
  let fileUrl = uploadToCloudflareR2(fileId, userName + "_" + getJstDateString());

  return fileUrl;
}

