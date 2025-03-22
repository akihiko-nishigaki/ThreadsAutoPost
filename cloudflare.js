function cloudflareMovieUpload(fileId)
{
  let userInfo = Session.getActiveUser();
  let userName = getLocalPartFromEmail(userInfo.getEmail());
  let movieUrl = bulkpostkun.uploadToCloudflareR2(fileId, userName + "_" + getJstDateString());

  return movieUrl;
}

