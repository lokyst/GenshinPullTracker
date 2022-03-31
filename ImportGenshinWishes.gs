function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Genshin Wishes")
    .addItem("Fetch Latest", "FetchAllWishes")
    .addToUi();
}

function FetchAllWishes(authKeySheet = "GenshinKey") {
  var BannerIdLookup ={
    "Novice": 100,
    "Permanent": 200,
    "Character Event": 301,
    "Weapon Event": 302
  }

  // Get the Authkey from the user
  var webURL = getURLFromSheet(authKeySheet);

  // Check validity of URL
  if (webURL.indexOf("https://webstatic-sea.hoyoverse.com/genshin/event/e20190909gacha/index.html?authkey_ver=1") < 0) {
    SpreadsheetApp.getUi().alert("Invalid URL format.")
    return;
  }

  var authkey = getAuthkeyFromURL(webURL);

  for (var banner in BannerIdLookup) {
    ImportWishesByBanner(authkey, banner, BannerIdLookup[banner]);
  }
}

function ImportWishesByBanner(authkey, sheetName, bannerId) {
  // Check sheet exists
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (mySheet == null) {
    mySheet = setupBannerSheet(sheetName);
  }
  
  // Read the last wish ID from the spreadsheet
  var c_lastWishID = "1612137599000000001";   // Default value
  var nSheetRows = mySheet.getLastRow();

  if (nSheetRows > 1) {
    c_lastWishID = mySheet.getRange(nSheetRows,5).getValue();
  }

  // Get wishes
  var wishes = getWishes(authkey, bannerId, c_lastWishID);

  // Check if any wishes were returned
  var nRows = wishes.length;
  if (nRows == 0) {
    // Pop up error message
    SpreadsheetApp.getUi().alert("No wishes found for banner: " + sheetName);
    return;
  }
  var nCols = wishes[0].length;

  // Write wishes back to spreadsheet in ascending chronological order
  mySheet.getRange(nSheetRows+1,1,nRows,nCols).setValues(wishes.reverse());

  // Setup 4* and 5* Pity counter columns 
  var str_formula_4_star = "=IF(ROW(R[0]C[0])<=2,1,IF(R[-1]C[-2]=4,1,R[-1]C[0]+1))";
  var str_formula_5_star = "=IF(ROW(R[0]C[0])<=2,1,IF(R[-1]C[-3]=5,1,R[-1]C[0]+1))";

  var destRange = mySheet.getRange(nSheetRows+1,6,mySheet.getLastRow()-nSheetRows,1);
  destRange.setFormulaR1C1(str_formula_4_star);
  destRange.offset(0,1).setFormulaR1C1(str_formula_5_star);

  // Setup conditional formatting
  applyConditionalFormatting(sheetName);

}

function getAuthkeyFromURL(myURL) {
  var authKeyStart = myURL.indexOf("authkey=")+"authkey=".length;
  var authKeyEnd = myURL.indexOf("&",authKeyStart);
  var authKey =  myURL.slice(authKeyStart, authKeyEnd);

  return authKey;
}

function constructURL(s_authKey, i_gacha_type, i_page, i_size, s_end_id) {
  var s_WishURL = "https://hk4e-api-os.mihoyo.com/event/gacha_info/api/getGachaLog?authkey_ver=1&sign_type=2&auth_appid=webview_gacha&init_type=301&lang=en&authkey="+ s_authKey +"&gacha_type="+ i_gacha_type + "&page=" + i_page + "&size=" + i_size + "&end_id=" + s_end_id;

  return s_WishURL;
}

function getWishes(authkey, i_gacha_type, c_lastWishID) {
  // Wish URL Parameters
  //var i_gacha_type = 200;
  var i_page = 1;
  var i_size = 10;
  var s_end_id = "0";

  // Loop through wish page URLs
  var wishes = [];  
  var count = 0;

  do {
    var s_WishURL= constructURL(authkey, i_gacha_type, i_page, i_size, s_end_id);
    var data = ImportJSON(s_WishURL, "/data/list", "noHeaders");
    var nRows = data.length;
    console.log(`Imported ${nRows} JSON rows`);
    if (nRows > 0) {
      var nCols = data[0].length;      
      var s_end_id = data[nRows-1][nCols-1];

      //wishes = wishes.concat(data);
      wishes = concatWishList(wishes, data, c_lastWishID);
      count++;
    }

  } while (s_end_id > c_lastWishID && nRows == i_size && count <= 100);

  return wishes;
}

function concatWishList(wishes, data, c_lastWishID) {
  var nRows = data.length;
  var nCols = data[0].length;
  

  for (var i=0; i < nRows; i++) {
    var s_lastDataListId = data[i][nCols-1];
    if (s_lastDataListId <= c_lastWishID) {
      break;
    }
    var newData = [data[i][4], data[i][5],data[i][7],Number(data[i][8]),data[i][9]];
    wishes = wishes.concat([newData]);
  }

  return wishes;
}

function applyConditionalFormatting(sheetName) {
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var destRange = mySheet.getRange(2,6,mySheet.getLastRow()-1,1);

  var rules = [];
  var str_color_4_star = "#b4a7d6";
  var str_formula_4_star = "=EQ(D2,4)";
  var str_color_5_star = "#ffd966";
  var str_formula_5_star = "=EQ(D2,5)";

  var str_color_min = "#57BB8A";
  var str_color_mid = "#FFD666";
  var str_color_max = "#E67C73";

  var rule_4_star = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(str_formula_4_star)
    .setBackground(str_color_4_star)
    .setRanges([destRange.offset(0,-4)])
    .build();
  rules.push(rule_4_star);

  var rule_5_star = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(str_formula_5_star)
    .setBackground(str_color_5_star)
    .setRanges([destRange.offset(0,-4)])
    .build();
  rules.push(rule_5_star);

  var rule_4_star_counter = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(str_color_min, SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue(str_color_mid, SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setGradientMaxpointWithValue(str_color_max, SpreadsheetApp.InterpolationType.NUMBER, "10")
    .setRanges([destRange])
    .build();
  rules.push(rule_4_star_counter);

  var rule_5_star_counter = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(str_color_min, SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue(str_color_mid, SpreadsheetApp.InterpolationType.NUMBER, "37")
    .setGradientMaxpointWithValue(str_color_max, SpreadsheetApp.InterpolationType.NUMBER, "82")
    .setRanges([destRange.offset(0,1)])
    .build();
  rules.push(rule_5_star_counter);

  // Clear existing rules and apply my rules
  mySheet.clearConditionalFormatRules();
  mySheet.setConditionalFormatRules(rules);
}

function setupBannerSheet(sheetName) {
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (mySheet == null) {
    mySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  mySheet.clear();

  var myRange = mySheet.getRange(1,1,1,7);

  // Create header row
  headerArray = [["Time", "Name", "Item Type", "Rank Type", "Id", "4*Pity", "5*Pity"]];
  myRange.setValues(headerArray);
  myRange.setFontWeight("bold");

  return mySheet;
}

function setupKeySheet(sheetName) {
  var keySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (keySheet == null) {
    keySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  keySheet.clear();

  var headerRowArray = [
    ["URL", "Enter Webstatic URL here"],
  ];

  var myRange = keySheet.getRange("A1");
  myRange = myRange.offset(0,0,1,2);
  myRange.setValues(headerRowArray);

  return keySheet;
}

function getURLFromSheet(mySheetname) {
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mySheetname);
  var myURL = "";

  if (mySheet == null) {
    // Create the sheet
    mySheet = setupKeySheet(mySheetname);

    // Have the user enter a URL
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter URL:");

    if (response.getSelectedButton() == ui.Button.OK) {
      myURL = response.getResponseText();

      // Save it to the key sheet
      mySheet.getRange("B1").setValue(myURL);
    } 
  }
  
  myURL = mySheet.getRange("B1").getValue() ;
  
  return myURL;
}




