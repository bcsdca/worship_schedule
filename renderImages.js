function renderImages() {
  const today = new Date();

  const comingSundayStr = getComingSunday();
  
  //const comingSundayMonth = comingSunday.getMonth();
  const comingSundayWeekOfMonth = getComingSundayWeekOfMonth();

  logMessage(
    getCallStackTrace() +
    `: Today: ${formatDate(today)}, Coming Sunday: ${comingSundayStr}, Week of Month: ${comingSundayWeekOfMonth}`
  );

  const images = {};

  // --- CEC Logo (Blob) ---
  try {
    const imageCECLogo = DriveApp
      .getFileById(cecLogoImageFileID)
      .getAs("image/png");

    if (imageCECLogo) {
      images.CEClogo = imageCECLogo;

      logMessage(
        getCallStackTrace() +
        `: Loaded CEC logo, size: ${imageCECLogo.getBytes().length} bytes`
      );
    }
  } catch (err) {
    logMessageError(
      `${getCallStackTrace()}: Failed to load CEC logo: ${err}`
    );
  }

  // --- Special Image Retrieval (from shared sheet) ---
  let imageSpecialBlob = null;

  try {
    const sheet = SpreadsheetApp
      .openById(weeklyShare)
      .getSheetByName("handoff");

    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];

      if (!rowDate) continue;
     
      if (formatDate(rowDate) === comingSundayStr) {
        const imageSpecialFileId = data[i][3];

        if (!imageSpecialFileId) {
          logMessageError(
            `${getCallStackTrace()}: Missing fileId in row ${i + 1}`
          );
          continue;
        }

        try {
          const file = DriveApp.getFileById(imageSpecialFileId);
          imageSpecialBlob = file.getBlob();

          logMessage(
            `${getCallStackTrace()}: Found special image for ${comingSundayStr}, FileId=${imageSpecialFileId}, Size=${imageSpecialBlob.getBytes().length} bytes`
          );

          break;
        } catch (err) {
          logMessageError(
            `${getCallStackTrace()}: Invalid file ID (${imageSpecialFileId}): ${err}`
          );
        }
      }
    }

  } catch (err) {
    logMessageError(
      `${getCallStackTrace()}: Error accessing shared sheet: ${err}`
    );
  }

  // --- Final assignment ---
  if (imageSpecialBlob) {
    images.special = imageSpecialBlob;
  } else {
    logMessageError(
      `${getCallStackTrace()}: No special image found for coming Sunday`
    );
  }

  return images;
}