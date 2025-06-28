
var html =
  '<body>' +
  '<h2> Test <img src = "cid:image"> </h2><br />' +
  '</body>'

function testGmailApp() {
  //https://drive.google.com/file/d/1Da9oiTFxiL-xr9tE1Vrs973Y7DPMybze/view?usp=sharing
  var ImageBlob = DriveApp
    .getFileById('1Da9oiTFxiL-xr9tE1Vrs973Y7DPMybze')
    .getBlob()
    .setName("ImageBlob");
  GmailApp.sendEmail(
    'shui.bill.chu@gmail.com',
    'test GmailApp',
    'test', {
    htmlBody: html,
    inlineImages: { image: ImageBlob }
  });
  console.log("sending test email")
}


