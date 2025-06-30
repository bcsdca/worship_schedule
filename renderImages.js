function renderImages() {
  // Calculate the date for the coming Sunday
  const today = new Date();
  var comingSunday = new Date();
  comingSunday.setDate(today.getDate() + (7 - today.getDay()));

  const comingSundayMonth = comingSunday.getMonth();

  var comingSundayWeekOfMonth = getComingSundayWeekOfMonth();

  logMessage(getCallStackTrace() + `Today: ${today}, Coming Sunday: ${comingSunday}, Month: ${comingSundayMonth}, Week of Month: ${comingSundayWeekOfMonth}`);

  // Default images
  const imageCECLogo = DriveApp.getFileById("1pCv6vjy8tbBnAHaUYiOjP9BdECqmWoAO").getAs("image/png");
  const imageDefault = DriveApp.getFileById("16CLqTLzFjkkcJc5iv1bri8ljy-Yu-ZBg").getAs("image/png");

  // Image file objects for special images categorized by month and week
  const images = {
    0: { // January
      1: DriveApp.getFileById("1qTAe3nK80r5lztSKEcBWVttqtJSz0fDd").getAs("image/gif"), // Happy New Year in Chinese
      2: DriveApp.getFileById("1zvQlTXGR8v4WONqJ4ZQG9S3JApnfc1Rj").getAs("image/gif"), // Happy New Year Jeremiah 29:11
      3: DriveApp.getFileById("1RRWuN3ncEwzGa3fAaj6ushu_qZqctsA-").getAs("image/gif"), // forward together
      4: DriveApp.getFileById("1NDUlDkCpLmM8iI1qKPz--Bd2XzVcs0-T").getAs("image/gif"), // John 1:14
      default: DriveApp.getFileById("1od5o18wJKLEugbgbI5scnQsWWJ9ok49N").getAs("image/gif") // Song of Solomon 2:12
    },
    1: { // February
      1: DriveApp.getFileById("1hw6hAHl1yqntUq_6v4QXduQhydisJuDl").getAs("image/gif"), // Pray Together
      2: DriveApp.getFileById("1bTqz4m0hq-_smnfwcUW9cvC8JT3vzTDe").getAs("image/gif"), // Chinese New Year
      3: DriveApp.getFileById("1bMEm5J_5ocFV7d3wUqHEuxhrQx_kyNmU").getAs("image/gif"), // Valentine’s Day
      4: DriveApp.getFileById("13r4Iky_fQgXDkzfiKDod6wYoHrCVn0iq").getAs("image/gif"), // Harvest is ripe, where are the workers
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    2: { // March
      1: DriveApp.getFileById("1EuQHiZD6bbf6QZTaQ5pWHPtfmmoreVkU").getAs("image/gif"), // Welcome March
      2: DriveApp.getFileById("12_dw2WZpzluOsL4MQkNUulP4LK3_PkLr").getAs("image/gif"), // Spring Forward
      3: DriveApp.getFileById("1eM2OZCf9kkHwyYJfCiV8g7X863FYdRoq").getAs("image/gif"), // God's Guidance
      4: DriveApp.getFileById("1vyAPy6mbYd2OuLJIsgzGgVZrKDC9ZwNh").getAs("image/gif"), // God's love march your way
      5: DriveApp.getFileById("104TL4FkXylhUvcz-AZgCWJuZthViKKDd").getAs("image/gif"), // Easter Sunday
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    3: { // April
      1: DriveApp.getFileById("17G-LKqet2PMCpz2E_P9aR3-Bz32nWioy").getAs("image/gif"), // Welcome April
      2: DriveApp.getFileById("1Cc4BUnsxyMqReyiYY_X-HY3VdDgrtlmv").getAs("image/gif"), // April Shower
      //2: DriveApp.getFileById("1sht70-D3KrqxPVDNAPbcNOWBX5EIpFL8").getAs("image/gif"), // April Shower
      3: DriveApp.getFileById("104TL4FkXylhUvcz-AZgCWJuZthViKKDd").getAs("image/gif"), // Easter Sunday
      //3: DriveApp.getFileById("1ZNlmHc9wkfajYiMkeApGfUXF9PTAP9v3").getAs("image/gif"), // Isaiah 43:5
      4: DriveApp.getFileById("1MTRX91wgbNbgoq2yf8BWgr2WIxCRqg5s").getAs("image/png"), // Psalm 91:2
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    4: { // May
      1: DriveApp.getFileById("1zXJcKQlht3jSxfmNURTiFPJNDDqd0Nha").getAs("image/gif"), // May God Bless You
      2: DriveApp.getFileById("1hAbkS1rrW_sNaML_5EcuBlRTrzJ5gYqU").getAs("image/gif"), // Happy Mother's Day
      3: DriveApp.getFileById("1p8MdT8SKBB5_VvFjaUJP4NkA1OhECkRK").getAs("image/gif"), // 1 Peter 5:7
      4: DriveApp.getFileById("1Q_z-32hsQmNZrDwYhxVuPAyFVy3-H6Tu").getAs("image/gif"), // Memorial Day
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    5: { // June
      1: DriveApp.getFileById("1_6xLq2dqhSnth71RqcsdKvfrMrXvpOtB").getAs("image/gif"), // Welcome June
      2: DriveApp.getFileById("1lcoL278fcehkwpYIMfR-150TzulZhy4q").getAs("image/gif"), // James 1:12
      3: DriveApp.getFileById("1g0j6ZkdT79JWEL0NPxJYmA6qdwHwzXXt").getAs("image/gif"), // Father's Day
      4: DriveApp.getFileById("1isBskX35M6CxjkW5gFL3YPacJahFmtqP").getAs("image/gif"), // Deuteronomy 31:8
      5: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif"), // Every Morning is New
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    6: { // July
      1: DriveApp.getFileById("1-VNFiBXoYsDctOeHqpR37KeCNhGIMR-i").getAs("image/png"), // July 4th with Bible Verse
      2: DriveApp.getFileById("1isBskX35M6CxjkW5gFL3YPacJahFmtqP").getAs("image/gif"), // Deuteronomy 31:8
      3: DriveApp.getFileById("1w0XYRq_u8_sar4j7wG5uw9l0cDU8yJW4").getAs("image/gif"), // Holy Spirit
      4: DriveApp.getFileById("11H8jdtMr6qMYTfgfUwu3VtQZdYgDkwWN").getAs("image/png"), // Psalm 145
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    7: { // August
      1: DriveApp.getFileById("15CV1dTg4d6kK3oDHqKkCdDfS5gLrNgji").getAs("image/png"), // Summer Break
      2: DriveApp.getFileById("10q6QpScsIOmZKB6aZ0PTWPT4yqWBEsfH").getAs("image/gif"), // Book of Jeremiah
      3: DriveApp.getFileById("100Y0Yp3gVTK7eTDZg0Asc83VqRtMgoZ0").getAs("image/gif"), // The Truth Shall Set You Free
      4: DriveApp.getFileById("1ChlkvOC5UGgwr4snzP7fNQkDa-JUnfVS").getAs("image/png"), // striving for unity 1cor1:10-17
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    8: { // September
      1: DriveApp.getFileById("1LeGlFpb8l6JaMnUuwd0zhdZyeRKCF1SZ").getAs("image/png"), // GENESIS 44
      2: DriveApp.getFileById("1maYlwYJteTag7_829D9QvHOzU6hcuKie").getAs("image/gif"), // Sept 11
      3: DriveApp.getFileById("1QmWcF8to_IrPuT2rHQH2n6YALI0KBnQN").getAs("image/gif"), // Hebrews1_jesus_is_superior
      4: DriveApp.getFileById("1ld8cH7T1FsvAAC4zElcgiLGt6VJUU3k3").getAs("image/png"), // The Church is the temple of God
      5: DriveApp.getFileById("19_k4jhS0ADNgCUNmThI9zMeFzJzaJxBH").getAs("image/png"), // The Church is the temple of God
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    9: { // October
      1: DriveApp.getFileById("1rh3jdDj2k7LA_xdVoKBPJOL6hebKOTOW").getAs("image/png"), // growing up in christ
      2: DriveApp.getFileById("1xmNlpvlGTx-naMAFw7F23PtntnV-gk18").getAs("image/png"), // hebrew3
      3: DriveApp.getFileById("1BF4uOBNICT8a3Lw9XWxjBReE8INWDIg-").getAs("image/gif"), // The church in trouble time
      4: DriveApp.getFileById("1vJDlnpf3XhIMf8ABdNGrzmo12j0gUB81").getAs("image/png"), // matt9:35-38
      default: DriveApp.getFileById("1p_apgtfIFByNBGLub3HBD3bJKfuofFgX").getAs("image/gif") // Every Morning is New
    },
    10: { // November
      //1: DriveApp.getFileById("1iiFnraUwvpnbRQZGrRVphgbm4vW4h1xa").getAs("image/png"), //trunk or treat 2024
      2: DriveApp.getFileById("11kpd-HsWRvKbDAAEINEckNhuKHAA_34-").getAs("image/png"), //acts15
      3: DriveApp.getFileById("16GoOEqevNg7updWe1NCUV03QHiyGJcJZ").getAs("image/gif"), // looking_at_the_blessings_of_God_kingdom
      4: DriveApp.getFileById("1lNplMtpajNDiRmoDF7JZmG3uNuqwJ5NU").getAs("image/gif"), //psalm103
      default: DriveApp.getFileById("1FT8vUPdy7GS0sGF06QBh6DF4drsePf8e").getAs("image/gif")
    },
    11: { // December
      1: DriveApp.getFileById("1y84e8-ObfrL09rIxzVvlglTYGoORWxAQ").getAs("image/gif"), // God's Molding
      2: DriveApp.getFileById("1y84e8-ObfrL09rIxzVvlglTYGoORWxAQ").getAs("image/gif"), // God's Molding
      3: DriveApp.getFileById("1JFF1j0D9a1iPHMM2qHbEMkjF4CVtIpFe").getAs("image/png"), // Immanuel_God_with_us
      4: DriveApp.getFileById("1Hg-KJsFxG9o6x7ufWniFenzm4PIp6_at").getAs("image/png"), // Epaphroditus' commitment Philippians 2: 25-30
      5: DriveApp.getFileById("1svr5FciRoGUeAwn_NgUlnzsxf5WGphw1").getAs("image/png"), // Andy working on laptop while on vacation
      default: DriveApp.getFileById("12fQgQpa4maZFvfWGiE87AwWTAVuhJywR").getAs("image/gif") // Every Morning is New
    }
  };

  let imageSpecial;

  // Select the correct image based on the coming Sunday’s month and week
  if (images[comingSundayMonth] && images[comingSundayMonth][comingSundayWeekOfMonth]) {
    imageSpecial = images[comingSundayMonth][comingSundayWeekOfMonth];
  } else if (images[comingSundayMonth] && images[comingSundayMonth].default) {
    imageSpecial = images[comingSundayMonth].default;
  }


  //return { "CEClogo": image_cecLogo, "special": image_special }
  logMessage(getCallStackTrace() + 'Size of imageCECLogo: ' + imageCECLogo.getBytes().length + ' bytes');
  logMessage(getCallStackTrace() + 'Size of imageSpecial: ' + imageSpecial.getBytes().length + ' bytes');

  return {
    CEClogo: imageCECLogo,
    special: imageSpecial
  };
}
