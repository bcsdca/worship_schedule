function sendEmailWithImages() {
  let image_cecLogo = DriveApp.getFileById("1Da9oiTFxiL-xr9tE1Vrs973Y7DPMybze").getAs("image/png");
  //let image_thanksGiving = DriveApp.getFileById("1itBHCzn0PAvzDSVtOvY0vxkG-ry6bwnL").getAs("image/png");
  //https://drive.google.com/file/d/1PB79a4P9e013jILcTmE9kGYS6zeppufN/view?usp=drive_link
  let image_thanksGiving = DriveApp.getFileById("1PB79a4P9e013jILcTmE9kGYS6zeppufN").getAs("image/gif");
  //let emailImages = { "CEClogo": image_cecLogo};
  let emailImages = { "CEClogo": image_cecLogo, "special": image_thanksGiving };
  let emailBody = HtmlService.createTemplateFromFile("template").evaluate().getContent();
  MailApp.sendEmail({
    to: "shui.bill.chu@gmail.com",
    subject: "Daily analytics report",
    htmlBody: emailBody,
    inlineImages: emailImages
  });
}

