/**
 * Google Apps Script サンプル: Cloudinary に画像／動画をアップロードする
 * Cloudinary 認証情報（ご自身のものに置き換えてください）
 */
const CLOUD_NAME = 'dxy2evcsj';
const API_KEY    = '291922288428472';
const API_SECRET = 'UpLZ1y_CrRR59OG5_KEYEvElUx4';

/**
 * Cloudinary に認証付きアップロードするための SHA-1 署名を生成します
 * @param {Object} params_to_sign - 署名に含めるパラメータのキー／値ペア
 * @return {string} - 16進文字列の署名
 */
function generateSignature(params_to_sign) {
  // パラメータをキーでソートして string_to_sign を作成
  const sortedKeys = Object.keys(params_to_sign).sort();
  const toSign = sortedKeys
    .map(key => `${key}=${params_to_sign[key]}`)
    .join('&');

  // HMAC-SHA1 で署名を生成
  const signatureBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_1,
    toSign,
    API_SECRET
  );

  // バイト配列を16進文字列に変換
  return signatureBytes
    .map(b => {
      const hex = (b < 0 ? b + 256 : b).toString(16);
      return hex.length === 1 ? '0' + hex : hex;
    })
    .join('');
}

/**
 * Google ドライブ上のファイル（画像または動画）を Cloudinary にアップロードします
 * @param {string} fileId - Google ドライブのファイル ID
 * @param {string} resourceType - 'image' または 'video'
 * @return {Object} - Cloudinary API レスポンスのパース結果
 */
function uploadFileToCloudinary(fileId, resourceType) {
  // ドライブから Blob を取得
  const blob = DriveApp.getFileById(fileId).getBlob();

  // Unix タイムスタンプを整数（秒）で取得
  const timestamp = Math.floor(Date.now() / 1000);

  // 署名対象のパラメータ

  // 署名対象のパラメータを準備
  const signingParams = { timestamp };
  // 画像の場合はデフォルトなので resource_type は署名／送信対象から除外
  if (resourceType && resourceType !== CONFIG.STRING_IMAGE) {
    signingParams.resource_type = resourceType;
  }
  const signature = generateSignature(signingParams);

  // フォームデータを構築
  const formData = {
    file: blob,
    api_key: API_KEY,
    timestamp: timestamp,
    signature: signature
  };

  if (resourceType && resourceType !== CONFIG.STRING_IMAGE) {
    formData.resource_type = resourceType;
  }

  // アップロード URL
  const url = `https://api.cloudinary.com/v1_1/${CLOUD_NAME}/${resourceType.toLowerCase()}/upload`;

  // ログ出力
  Logger.log('url: ' + url);

  // リクエスト実行
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: formData,
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  Logger.log(result);
  return result;
}

/**
 * Google ドライブの画像ファイルを Cloudinary にアップロードします
 * @param {string} fileId - Google ドライブのファイル ID
 */
function uploadImageToCloudinary(fileId) {
  const result = uploadFileToCloudinary(fileId, CONFIG.STRING_IMAGE);
  // アップロード結果から画像 URL をログ出力
  Logger.log('画像の URL: ' + result.secure_url);
}

/**
 * Google ドライブの動画ファイルを Cloudinary にアップロードします
 * @param {string} fileId - Google ドライブのファイル ID
 */
function uploadVideo(fileId) {
  const result = uploadFileToCloudinary(fileId, CONFIG.STRING_VIDEO);
  // アップロード結果から動画 URL をログ出力
  Logger.log('動画の URL: ' + result.secure_url);
}
