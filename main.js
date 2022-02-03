const TO_FIELD = "B2"
const SUBJECT_FIELD = "B3"
const CC_FIELD = "B4"
const BCC_FIELD = "B5"
const BODY_FIELD = "A6"
const DATA_RANGE = "G1:K7"

const SEND_MAIL = false

function myFunction() {
    var sheet = SpreadsheetApp.getActiveSheet();
    const rangeName = 'A2:A3';
    // Get the values from the spreadsheet using spreadsheetId and range.
    const values2 = sheet.getRange("a1").values;
    var cell = SpreadsheetApp.getActive().getRange('A6').getRichTextValue();
    Logger.log(cell);

    //Logger.log(cell.getTextStyle().isBold())
    var emailTemplateTo = sheet.getRange(TO_FIELD).getValue();
    var emailTemplateSubject = sheet.getRange(SUBJECT_FIELD).getValue();
    var emailTemplateCC = sheet.getRange(CC_FIELD).getValue();
    var emailTemplateBCC = sheet.getRange(BCC_FIELD).getValue();
    var emailTemplateBodyRichValue = sheet.getRange(BODY_FIELD).getRichTextValue();
    var inputData = sheet.getRange(DATA_RANGE).getValues();
    var inputDataHeaders = inputData[0];
    inputData.shift() //removes headers from inputData
    Logger.log(inputDataHeaders)
    Logger.log(inputData)
    Logger.log(emailTemplateBodyRichValue)
    //TODO: remove blank cells from data


    var headersList = SpreadsheetApp.getActive().getRange('G1:I1').getValues()[0];
    var values2DList = SpreadsheetApp.getActive().getRange('G2:I7').getValues();
    Logger.log(inputData.length)
    for (var i = 0; i < inputData.length; i++){

      var dict = {};
      for (var j = 0; j < inputDataHeaders.length; j++ ){
        dict[inputDataHeaders[j]] = inputData[i][j];

      }
      var emailArgs = {};
      emailArgs["to"] = parseStringTemplate(emailTemplateTo, dict)
      emailArgs["subject"] = parseStringTemplate(emailTemplateSubject, dict)
      emailArgs["cc"] = parseStringTemplate(emailTemplateCC, dict)
      emailArgs["bcc"] = parseStringTemplate(emailTemplateBCC, dict)
      emailArgs["htmlBody"] = parseStringTemplate(richTextToHTML(emailTemplateBodyRichValue),dict)
      Logger.log(emailArgs["to"])
      Logger.log(emailArgs["htmlBody"]);
      sendEmail({to: emailArgs["to"], name: emailArgs["name"], subject: emailArgs["subject"], htmlBody: emailArgs["htmlBody"]})


    }
}

function parseCSV (csv, dict){
  csvList = csv.split(",");
  for (i = 0; i < csvList.length; i++){
    i = i;
  }
}

function sendEmail(args){
  var test = true;
  if (SEND_MAIL){
    MailApp.sendEmail({to: args["to"],name: args["name"],subject: args["subject"],htmlBody: args["htmlBody"]});
    test = false
  }
  if(test){test = "TEST ";}else{test = "";}
  Logger.log(test+"EMAIL SENT\nname: " + args["name"] + "\nto: " + args["to"])
}

function richTextToHTML(richTextObject){
    //documentation https://developers.google.com/apps-script/reference/spreadsheet/rich-text-value
    richTextList = richTextObject.getRuns();
    HTMLoutput = "";
    for (var i = 0; i < richTextList.length; i++){
      richTextRun = richTextList[i];
      textStyle = richTextRun.getTextStyle();
      text = richTextRun.getText();
      if(richTextRun.getLinkUrl()){text = "<a href='" + richTextRun.getLinkUrl() + "'>" + text + "</a>"}
      if(textStyle.isBold()){text = "<strong>" + text +"</strong>";}
      if(textStyle.isItalic()){text = "<i>" + text +"</i>";}
      if(textStyle.isStrikethrough()){text = "<strike>" + text +"</strike>";}
      if(textStyle.isUnderline()){text = "<u>" + text +"</u>";}
      if(text.includes("\n")){
        Logger.log("yeeeeeeeeeee")
      text = text.replaceAll("\n", "<br>")
      }
      if(textStyle.getForegroundColor() && textStyle.getForegroundColor() != "#000000"){
        //gets hex color, can't use "textStyle.getForegroundColor()" cause returns things like "ACCENT2" which don't work outside of the sheet
        hexColor = textStyle.getForegroundColorObject().asRgbColor().asHexString()
        //Logger.log(hexColor)
        text = "<span style='color:" + hexColor + "'>" + text +"</span>";
        }
        //TODO: remove styling if it is a link
        //TODO: make links work
      HTMLoutput += text
    }
    //Logger.log(HTMLoutput)
    return HTMLoutput;
}

//modified function from pekaaw on StackOverflow https://stackoverflow.com/a/59084440
function parseStringTemplate(str, obj) {
    //let parts = str.split(/\$\{(?!\d)[\wæøåÆØÅ]*\}/);
    let parts = str.split(/{(?!\d)[\wæøåÆØÅ]*\}/);
    let args = str.match(/[^{\}]+(?=})/g) || [];
    let parameters = args.map(argument => obj[argument] || (obj[argument] === undefined ? "" : obj[argument]));
    return String.raw({ raw: parts }, ...parameters);
}
